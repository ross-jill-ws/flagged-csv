# Flagged-CSV Implementation Notes

## Overview
This document contains implementation notes and lessons learned from developing the flagged-csv package, which converts XLSX files to CSV while preserving visual formatting as inline flags.

## Critical Implementation Details

### 1. Color Handling

#### Foreground and Background Colors
- **Syntax**: 
  - Background colors: `{#RRGGBB}` (backward-compatible) or `{bc:#RRGGBB}`
  - Foreground colors: `{fc:#RRGGBB}`
- **Important**: Only extract foreground colors for cells that have content (`value is not None`)

#### Black Text Contrast Logic
- **Issue**: Black text (`{fc:#000000}`) was appearing unnecessarily on cells without backgrounds, creating visual noise
- **Solution**: Only include black foreground color when there's a background color for contrast
- **Implementation**: Check if cell has background before including black text flag
```python
# Only include black foreground color if there's also a background color (for contrast)
if color_to_check.upper() == '000000':
    if not has_background:
        # Need to check for background if we haven't already
        bg_color = self._extract_cell_bg_color(cell, file_path)
        has_background = bg_color is not None
    # Only include black text if there's a background
    if has_background:
        formatting_parts.append(f"{{fc:{fg_color_hex}}}")
```

#### Default Ignore Colors
- **Background colors**: Ignore `#FFFFFF` (white) by default
- **Foreground colors**: Ignore `#000000` (black) by default
- **Rationale**: These are the most common default colors that don't need explicit flagging
- **CLI Options**:
  - `--ignore-colors`: Applies both defaults
  - `--ignore-bg-colors`: Custom background ignore list (default: `#FFFFFF`)
  - `--ignore-fg-colors`: Custom foreground ignore list (default: `#000000`)

### 2. Indexed Color Handling

#### Critical Issue: Hardcoded Indexed Colors
- **Problem**: Initially used a hardcoded mapping for indexed colors, which was incorrect
- **Example**: Indexed color 9 was defaulting to `#000000` (black) instead of `#FFFFFF` (white)
- **Root Cause**: Excel files can have custom indexed color palettes; hardcoding is unreliable

#### Correct Solution: Dynamic Color Palette Reading
- **Primary Source**: Read from `workbook._colors` attribute when available
- **Fallback**: Use `openpyxl.styles.colors.COLOR_INDEX` for standard colors
- **Implementation**:
```python
# First try to get from workbook's custom colors if available
wb = cell.parent.parent  # cell -> worksheet -> workbook
if hasattr(wb, '_colors') and wb._colors and idx < len(wb._colors):
    rgb = wb._colors[idx]
    # Remove alpha channel if present (first 2 chars)
    if len(rgb) == 8 and rgb.startswith('00'):
        return f"#{rgb[2:]}"
        
# Fall back to openpyxl's default COLOR_INDEX
from openpyxl.styles.colors import COLOR_INDEX
if idx < len(COLOR_INDEX):
    rgb = COLOR_INDEX[idx]
```

### 3. Testing Strategy

#### Color Testing
- Always test with actual Excel files that use indexed colors
- Verify that white backgrounds (`#FFFFFF`) are properly ignored by default
- Test black text contrast logic with various background combinations
- Create test cases for:
  - White text on dark backgrounds
  - Black text on light backgrounds (should be ignored by default)
  - Black text without background (should never appear)
  - Colored text without background (should appear)

#### Indexed Color Testing
- Create test files with indexed colors to ensure proper extraction
- Verify against actual Excel behavior, not assumptions
- Test both custom and standard indexed color palettes

### 4. Common Pitfalls to Avoid

1. **Never hardcode indexed color mappings** - Always read from the workbook or use openpyxl's defaults
2. **Don't extract foreground colors for empty cells** - Check `value is not None` first
3. **Remember alpha channels** - Excel color strings often have 8 characters (AARRGGBB), strip the alpha
4. **Test with real Excel files** - Synthetic test files may not reveal all edge cases
5. **Consider contrast** - Black text on white background is redundant; implement smart defaults

### 5. CLI Design Principles

- **Granular control**: Provide separate options for foreground/background colors
- **Smart defaults**: Use sensible ignore lists (#FFFFFF for bg, #000000 for fg)
- **Backward compatibility**: Maintain old syntax while adding new features
- **Master flags**: `--include-colors` as a convenience flag for both fg and bg

### 6. Version Management

When fixing bugs related to color handling:
1. Always bump the patch version for bug fixes
2. Update CHANGELOG.md with clear description of what was fixed
3. Include test cases that verify the fix
4. Document any behavior changes that might affect existing users

## Key Lessons Learned

1. **Excel's complexity**: Excel has multiple ways to represent colors (theme, indexed, RGB), and each needs proper handling
2. **Trust the library**: openpyxl already handles much of the complexity; use its built-in features
3. **User feedback is crucial**: The black text contrast issue was only discovered through actual usage
4. **Test with production data**: Real Excel files often have edge cases not present in synthetic tests
5. **Document defaults clearly**: Users need to understand what colors are ignored by default and why