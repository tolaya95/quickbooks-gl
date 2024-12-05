# quickbooks-gl
let
    Source = Excel.Workbook(File.Contents("C:\Users\tolay\OneDrive\Desktop\Quickwash Data Connections\QuickBooks Data Tables.xlsm"), null, true),
    #"QuickBooks General Ledger (2)_Sheet" = Source{[Item="QuickBooks General Ledger (2)",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"QuickBooks General Ledger (2)_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type text}, {"Transaction Type", type text}, {"Name", type text}, {"Memo/Description", type text}, {"Split", type text}, {"Amount", type number}, {"Account", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Transaction Type] <> null)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Account"}, qw_expense_account_list, {"Account"}, "QW Expense Account List", JoinKind.LeftOuter),
    #"Expanded QW Expense Account List" = Table.ExpandTableColumn(#"Merged Queries", "QW Expense Account List", {"Grouped Category"}, {"QW Expense Account List.Grouped Category"}),
    #"Filtered Rows1" = Table.SelectRows(#"Expanded QW Expense Account List", each ([QW Expense Account List.Grouped Category] <> null and [QW Expense Account List.Grouped Category] <> "N/A")),
    #"Changed Type1" = Table.TransformColumnTypes(#"Filtered Rows1",{{"Date", type date}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"Date", Order.Descending}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Sorted Rows",{{"Amount", Currency.Type}})
in
    #"Changed Type2"
