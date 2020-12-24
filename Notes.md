# Notes

## Search
- If Forestroot is the base, then searc must be against GC (Global Catalog) `-gc`.
- Test searches against results with spaces: use `scientifc*` as search key
- `Query * -filter (attribute=<keyword)` keyword can not be in quotes
- Using dsquery computer `-name` with `-desc *` only returns those computers with a description.
- Using DSQUERY * -filter managedBy, must be the complete DN --* doesn't work
- Can use this [epochconverter website](https://www.epochconverter.com/ldap) to convert NT time both ways.