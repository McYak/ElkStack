[xml]$DATA=get-content '.\sheet1.xml' 
foreach ($row in $DATA.worksheet.sheetData.row){
    foreach ($column in $row.c[5]){
        Write-Host $column.v
        }
        
    }
