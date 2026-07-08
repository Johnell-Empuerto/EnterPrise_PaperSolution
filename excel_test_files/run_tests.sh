#!/bin/bash
# Phase 1.2 - Comprehensive Test Runner
# Tests each Excel workbook against the API and documents results.

API_URL="http://localhost:5090/api/excel/upload"
TEST_DIR="excel_test_files/test_files_output"
REPORT_FILE="excel_test_files/test_results.csv"

echo "Phase 1.2 - Excel Capture API Test Results"
echo "=========================================="
echo ""
echo "API Endpoint: $API_URL"
echo "Test Files:   $TEST_DIR"
echo ""

# CSV header
echo "Test File,Status,HTTP Code,Success,Message,Image URL,File Size (KB),Time (s)" > "$REPORT_FILE"

total=0
passed=0
failed=0
error_expected=0
error_got=0

for file in "$TEST_DIR"/*.xlsx; do
    filename=$(basename "$file")
    filesize=$(du -k "$file" | cut -f1)
    total=$((total + 1))

    echo -n "  Testing: $filename ... "

    # Time the request
    start_time=$(date +%s.%N)
    
    # Make the API call
    response=$(curl -s -w "\n%{http_code}" -X POST \
        -F "file=@$file" \
        "$API_URL" 2>&1)
    
    end_time=$(date +%s.%N)
    duration=$(echo "$end_time - $start_time" | bc 2>/dev/null || echo "0")

    # Split response and HTTP code
    http_code=$(echo "$response" | tail -1)
    body=$(echo "$response" | head -n -1)

    # Parse JSON response
    success=$(echo "$body" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('success',''))" 2>/dev/null)
    message=$(echo "$body" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('message',''))" 2>/dev/null)
    image_url=$(echo "$body" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('data',{}).get('imageUrl',''))" 2>/dev/null)

    # Determine pass/fail
    # File 20 should return error (no print area)
    if [[ "$filename" == "20_empty_print_area"* ]]; then
        if [[ "$http_code" == "400" ]]; then
            echo "PASS (expected 400)"
            error_got=$((error_got + 1))
            status="PASS (expected)"
        else
            echo "FAIL - expected 400 but got $http_code"
            failed=$((failed + 1))
            status="FAIL"
        fi
        error_expected=$((error_expected + 1))
    elif [[ "$http_code" == "200" && "$success" == "True" ]]; then
        echo "PASS (HTTP $http_code, $duration s)"
        passed=$((passed + 1))
        status="PASS"
        
        # Download and check the preview image
        if [[ -n "$image_url" ]]; then
            preview_url="http://localhost:5090$image_url"
            preview_file="/tmp/preview_${filename%.xlsx}.png"
            curl -s -o "$preview_file" "$preview_url" 2>/dev/null
            if [[ -f "$preview_file" ]]; then
                png_size=$(stat -c%s "$preview_file" 2>/dev/null || wc -c < "$preview_file" 2>/dev/null)
                echo "         Preview PNG: ${png_size} bytes"
            fi
        fi
    else
        echo "FAIL (HTTP $http_code, Success=$success)"
        failed=$((failed + 1))
        status="FAIL"
    fi

    # Log to CSV
    echo "$filename,$status,$http_code,$success,\"$message\",$image_url,$filesize,$duration" >> "$REPORT_FILE"
    echo ""
done

# Summary
echo ""
echo "=========================================="
echo "Test Summary"
echo "=========================================="
echo "Total tests:  $total"
echo "Passed:       $passed"
echo "Failed:       $failed"
echo "Error tests:  $error_expected (expected $error_got)"
echo ""
echo "Detailed results: $REPORT_FILE"
