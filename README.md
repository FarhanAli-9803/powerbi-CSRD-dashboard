# CSRD Sustainability Analysis Dashboard 🌍

A Power BI dashboard project analyzing **Sustainability Data Points** aligned with the **European Sustainability Reporting Standards (ESRS)**.  
This project combines Power BI visuals, DAX logic, and Excel macros for a powerful reporting tool.

---

## 📊 Dashboard Features

- Clean, executive-ready design for sustainability stakeholders
- Slicers and filters for dynamic analysis by company, scope, or ESG category
- Summary visuals, KPIs, and breakdowns across ESRS themes
- Color-coded risk levels (via Excel VBA)

---

## 🖼 Dashboard Preview

![Page 1](image/CSRD%20Dashboard%20Page%201.png)
![Page 2](image/CSRD%20Dashboard%20Page%202.png)
![Page 3](image/CSRD%20Dashboard%20Page%203.png)
![Page 4](image/CSRD%20Dashboard%20Page%204.png)
![Page 5](image/CSRD%20Dashboard%20Page%205.png)

---

## 📥 Data Source

📁 [`Case - DE.xlsm`](excel-source/Case%20-%20DE.xlsm)

This Excel file contains:

- Raw CSRD indicators and company data
- VBA macros to extract:
  - `FontColor`
  - `CellColor`
  - `Classification` based on color
- Used as the main Power BI data source

### 🔧 Key VBA Functions

```vba
Function GetFontColor(cell As Range) As Long
    GetFontColor = cell.Font.Color
End Function

Function GetCellColor(cell As Range) As Long
    GetCellColor = cell.Interior.Color
End Function
