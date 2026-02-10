# Kusto to Markdown Converter

A Windows utility that converts Kusto Explorer clipboard data into clean Markdown — ready to paste into GitHub issues, wikis, or documentation.

## What it does

1. **Reads** the HTML clipboard data that Kusto Explorer places on the Windows clipboard when you copy query results.
2. **Extracts** the KQL query, result table, cluster URL, and deep-links.
3. **Converts** everything into a well-formatted Markdown document.
4. **Copies** the Markdown back to your clipboard.

## Output includes

- **Query section** with the KQL in a fenced code block (` ```kql `)
- **Cluster URL** and **"Open in"** deep-links to Kusto Explorer / Web Explorer
- **Results table** formatted as a Markdown table with auto-linkified URLs

## Requirements

- **Windows** (uses the Win32 clipboard API via `ctypes`)
- **Python 3.10+**
- **pyperclip** — auto-installed on first run if missing

## Usage

1. In **Kusto Explorer**, run a query and copy the results (`Ctrl+C`).
2. Run the script:

   ```powershell
   python kusto2md.py
   ```

3. The Markdown is now in your clipboard — paste it wherever you need.

### Fallback mode

If no Kusto HTML data is found on the clipboard, the script falls back to treating the clipboard content as **tab-separated values** and converts that into a Markdown table.

## Example output

```markdown
### Query

> **Cluster:** https://mycluster.kusto.windows.net
> **Open in:** [Kusto Explorer](https://...) | [Web](https://...)

​```kql
StormEvents
| summarize count() by State
| top 5 by count_
​```

### Results

| State          | count_ |
| -------------- | ------ |
| TEXAS          | 4701   |
| KANSAS         | 3166   |
| OKLAHOMA       | 2690   |
| MISSOURI       | 2016   |
| GEORGIA        | 1983   |
```

## License

MIT
