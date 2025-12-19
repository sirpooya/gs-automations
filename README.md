# Google Sheets Invoice Generator

A Google Apps Script that generates and displays an invoice summary from a "Report" sheet with a custom menu item.

## Features

- Custom menu item "📜 Invoice" with "Show Invoice" submenu
- Automatically extracts pricing and team data from the Report sheet
- Displays formatted invoice summary in a modal dialog
- Supports Persian/Farsi text formatting

## Setup

1. Open your Google Sheet
2. Go to **Extensions** → **Apps Script**
3. Copy and paste the code from `invoice-maker.gs` into the script editor
4. Save the project (Ctrl+S or Cmd+S)
5. Refresh your Google Sheet
6. The "📜 Invoice" menu will appear in the menu bar

## Requirements

Your Google Sheet must have a sheet named **"Report"** with the following structure:

### Column Headers (Row 1)
- `Full` - Full seat price
- `Dev` - Developer seat price
- `Collab` - Collaboration seat price
- `Team` - Team names
- `Cost` - Cost values

### Data Structure
- **Row 2**: Contains price values for Full, Dev, and Collab columns
- **Rows 3+**: Team data (name, seat counts)
- **"Subtotal" row**: Contains totals in Team column
- **"Prorated Costs" row**: Contains prorated cost value
- **"Total" row**: Contains final total cost

## Usage

1. Click on **📜 Invoice** in the menu bar
2. Select **Show Invoice**
3. A modal dialog will display the invoice summary with:
   - Price per seat for each type
   - Team breakdown with seat counts
   - Total seat counts
   - Final cost calculation

## Output Format

The invoice displays:
- Price per seat: Full, Dev, Collab
- Each team's seat allocation
- Total seats summary
- Cost breakdown (subtotal + prorated costs = total)

## Error Handling

The script includes error handling for:
- Missing "Report" sheet
- Missing required columns
- Missing "Subtotal" row
- Missing "Prorated Costs" or "Total" rows (defaults to 0)

