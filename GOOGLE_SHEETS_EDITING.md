# Google Sheets Orders Editing Feature

This feature allows you to edit certain fields in the Google Sheets **Orders** worksheet and have those changes persist when you run the Minifig Profit Tool again.

## How It Works

When the tool runs, it:
1. Reads any existing user edits from the Orders worksheet
2. Processes the latest order data from your XML/CSV files
3. Merges your edits with the fresh data
4. Updates the worksheet with the combined data

## Editable Fields

You can edit the following fields in the Orders worksheet, and your changes will be preserved:

- **Shipping** - Add shipping costs for orders
- **Add Chrg 1** - Additional charges (fees, taxes, etc.)  
- **Total Lots** - Number of different item types in the order
- **Total Items** - Total quantity of all items in the order
- **Tracking No** - Shipping tracking numbers

## Usage Instructions

1. **Run the tool normally** to populate the Orders worksheet with your BrickLink order data
2. **Edit the desired fields** directly in Google Sheets:
   - Open your "Minifig Profit Tool" spreadsheet
   - Go to the "Orders" tab
   - Edit any of the supported fields (Shipping, Add Chrg 1, Total Lots, Total Items, Tracking No)
   - Save your changes in Google Sheets
3. **Run the tool again** - your edits will be preserved and merged with any new order data

## Example Workflow

```
Initial state (from BrickLink data):
Order ID: 123456, Seller: TestSeller, Shipping: [empty], Tracking No: [empty]

User edits in Google Sheets:
Order ID: 123456, Seller: TestSeller, Shipping: $5.99, Tracking No: 1Z999AA1234567890

After next tool run:
Order ID: 123456, Seller: TestSeller, Shipping: $5.99, Tracking No: 1Z999AA1234567890
(Edits preserved!)
```

## Important Notes

- **Supported fields only**: Only the fields listed above will preserve edits. Other fields are regenerated from your BrickLink data
- **No validation**: The tool doesn't validate your edits - make sure shipping costs, tracking numbers, etc. are correct
- **Order identification**: Edits are tied to specific orders and items by Order ID and Item Number
- **Backup recommended**: Always keep backups of important data, as with any automated tool

## Troubleshooting

- **Edits not preserved**: Make sure you're editing the supported fields and that your Google Sheets file is saved
- **Formatting issues**: Numeric fields like Shipping should contain valid numbers (e.g., "5.99" not "$5.99" - the tool will format currency display)
- **Permission errors**: Ensure the service account has proper access to your Google Sheets

## Technical Details

The feature works by:
1. Calling `read_orders_sheet_edits()` to capture existing user edits before clearing the worksheet
2. Using `update_orders_sheet()` to merge preserved edits with fresh order data
3. Applying edits based on (Order ID, Item Number) matching keys

This preserves the existing workflow while adding edit persistence functionality.