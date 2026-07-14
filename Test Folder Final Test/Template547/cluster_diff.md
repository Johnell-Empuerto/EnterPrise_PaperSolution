# def_cluster Diff (Address-Matched)

## DB Clusters vs Generated Clusters

| DB Addr | Gen Addr | DB Left | Gen Left | DB Top | Gen Top | Addr Match |
|---------|----------|---------|----------|--------|---------|------------|
| $I$6:$M$6 | $I$6:$M$6 | 0.5294118 | 0.5294118 ✓ | 0.1654546 | 0.1654546 ✓ | ✓ |
| $I$7:$M$7 | $I$7:$M$7 | 0.5294118 | 0.5294118 ✓ | 0.1877273 | 0.1821213 ≠ | ✓ |
| $I$8:$M$8 | $I$8:$M$8 | 0.5294118 | 0.5294118 ✓ | 0.2104545 | 0.1987879 ≠ | ✓ |
| $I$9:$M$9 | $I$9:$M$9 | 0.5294118 | 0.5294118 ✓ | 0.2336364 | 0.2154546 ≠ | ✓ |
| $I$10:$M$10 | $I$10:$M$10 | 0.5294118 | 0.5294118 ✓ | 0.2568182 | 0.2321213 ≠ | ✓ |

## Extra Generated Clusters (not in DB) — 100 total

- $E$1
- $F$1
- $G$1
- $B$2:$M$2
- $L$3:$M$3
- $B$4:$L$4
- $B$5:$L$5
- $B$6:$G$12
- $H$6
- $R$6
- $H$7
- $H$8
- $H$9
- $H$10
- $H$11
- $I$11
- $J$11
- $K$11
- $L$11:$M$11
- $H$12
- ... and 80 more

## Notes
- Addr Match: ✓ = cluster found in generated XML by cell address
- ≠ = addresses match but position values differ
- ✗ = cluster address not found in generated XML
- Extra clusters in generated XML are expected because the legacy DB
  only stores user-configured clusters (input fields), while the engine
  generates clusters for ALL merged cells and standalone cells.
