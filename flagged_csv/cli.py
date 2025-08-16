"""
Command-line interface for flagged-csv converter.
"""

import click
from pathlib import Path
from .converter import XlsxConverter, XlsxConverterConfig


@click.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('-t', '--tab-name', required=True, help='Sheet/tab name to convert')
@click.option('-o', '--output', type=click.Path(), help='Output file (default: stdout)')
@click.option('--format', type=click.Choice(['csv', 'html', 'markdown']), default='csv', 
              help='Output format (default: csv)')
@click.option('--include-colors', is_flag=True, help='Include both foreground and background colors')
@click.option('--include-bg-colors', is_flag=True, help='Include background colors as {#RRGGBB} flags')
@click.option('--include-fg-colors', is_flag=True, help='Include foreground colors as {fc:#RRGGBB} flags')
@click.option('--signal-merge', is_flag=True, help='Include merged cell info as {MG:XXXXXX} flags')
@click.option('--preserve-formats', is_flag=True, help='Preserve cell formatting (e.g., $500 vs 500)')
@click.option('--ignore-colors', type=str, help='Ignore colors for both fg and bg (applies defaults: #FFFFFF for bg, #000000 for fg)')
@click.option('--ignore-bg-colors', type=str, help='Ignore specific background colors (default: #FFFFFF)')
@click.option('--ignore-fg-colors', type=str, help='Ignore specific foreground colors (default: #000000)')
@click.option('--keep-empty-lines', is_flag=True, help='Preserve empty rows to maintain original row positions')
@click.option('--add-location', is_flag=True, help='Add location coordinates {l:A5} to non-empty cells')
@click.option('--max-rows', type=int, default=300, help='Maximum rows to process (default: 300)')
@click.option('--max-columns', type=int, default=100, help='Maximum columns to process (default: 100)')
@click.option('--no-header', is_flag=True, help='Exclude header row from output')
@click.option('--keep-na', is_flag=True, help='Keep NA values instead of converting to empty strings')
def main(input_file, tab_name, output, format, include_colors, include_bg_colors,
         include_fg_colors, signal_merge, preserve_formats, ignore_colors, 
         ignore_bg_colors, ignore_fg_colors, keep_empty_lines, add_location, 
         max_rows, max_columns, no_header, keep_na):
    """
    Convert XLSX files to CSV with visual formatting preserved as inline flags.
    
    Examples:
    
        # Basic conversion
        flagged-csv data.xlsx -t Sheet1
        
        # Include colors and merge info
        flagged-csv data.xlsx -t Sheet1 --include-colors --signal-merge
        
        # Save to file with formatting preserved
        flagged-csv data.xlsx -t Sheet1 --preserve-formats -o output.csv
        
        # Ignore white background
        flagged-csv data.xlsx -t Sheet1 --include-colors --ignore-colors "#FFFFFF"
    """
    try:
        # Create converter with config
        config = XlsxConverterConfig(
            header=not no_header,
            keep_default_na=keep_na,
            keep_empty_lines=keep_empty_lines,
            add_location=add_location
        )
        converter = XlsxConverter(config)
        
        # Convert the file
        result = converter.convert_to_csv(
            input_file,
            tab_name=tab_name,
            output_format=format,
            include_colors=include_colors,
            include_bg_colors=include_bg_colors,
            include_fg_colors=include_fg_colors,
            signal_merge=signal_merge,
            preserve_formats=preserve_formats,
            ignore_colors=ignore_colors,
            ignore_bg_colors=ignore_bg_colors,
            ignore_fg_colors=ignore_fg_colors,
            keep_empty_lines=keep_empty_lines,
            add_location=add_location,
            max_rows=max_rows,
            max_columns=max_columns
        )
        
        # Output result
        if output:
            Path(output).write_text(result, encoding='utf-8')
            click.echo(f"Converted to {output}")
        else:
            click.echo(result, nl=False)
            
    except FileNotFoundError as e:
        click.echo(f"Error: {e}", err=True)
        raise click.Abort()
    except ValueError as e:
        click.echo(f"Error: {e}", err=True)
        raise click.Abort()
    except Exception as e:
        click.echo(f"Unexpected error: {e}", err=True)
        raise click.Abort()


if __name__ == '__main__':
    main()