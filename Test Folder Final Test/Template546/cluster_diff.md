# def_cluster Diff (Address-Matched)

## DB Clusters vs Generated Clusters

| DB Addr | Gen Addr | DB Left | Gen Left | DB Top | Gen Top | Addr Match |
|---------|----------|---------|----------|--------|---------|------------|
| $A$1:$B$2 | $A$1:$B$2 | 0.3364706 | 0.3364706 ✓ | 0.3845454 | 0.3845454 ✓ | ✓ |
| $C$1:$D$2 | $C$1:$D$2 | 0.5 | 0.4933334 ≠ | 0.3845454 | 0.3845454 ✓ | ✓ |
| $A$3:$D$4 | $A$3:$D$4 | 0.3364706 | 0.3364706 ✓ | 0.4231818 | 0.420909 ≠ | ✓ |
| $A$6:$D$7 | $A$6:$D$7 | 0.3364706 | 0.3364706 ✓ | 0.4809091 | 0.4754545 ≠ | ✓ |
| $A$9:$D$10 | $A$9:$D$10 | 0.3364706 | 0.3364706 ✓ | 0.5386364 | 0.53 ≠ | ✓ |
| $A$12 | $A$12 | 0.3364706 | 0.3364706 ✓ | 0.5963637 | 0.5845454 ≠ | ✓ |

## Notes
- Addr Match: ✓ = cluster found in generated XML by cell address
- ≠ = addresses match but position values differ
- ✗ = cluster address not found in generated XML
- Extra clusters in generated XML are expected because the legacy DB
  only stores user-configured clusters (input fields), while the engine
  generates clusters for ALL merged cells and standalone cells.
