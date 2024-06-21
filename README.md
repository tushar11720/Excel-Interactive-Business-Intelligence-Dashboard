# Interactive Business Intelligence Dashboard

## Overview

This project involves the creation of an interactive business intelligence dashboard in Excel. The dashboard maximizes data visualization and analysis capabilities, providing insights into customer behavior, product profitability, and store performance. It includes VBA scripting to automate repetitive tasks, enhancing operational efficiency and decision-making processes.

## Features

- **Interactive Dashboard:** Comprehensive visualization of key business metrics.
- **Customer Analysis:** Insights into profit generated from male and female customers, with a breakdown of average spending by age groups.
- **Profitability Trends:** Analysis of profit trends over time, including month-over-month growth rates.
- **Weekday Profitability:** Identification of the most profitable days of the week.
- **Product Analysis:** Details of top-selling and most profitable products, along with return and refund rates.
- **Store Performance:** Comparison of revenue vs. target for each store and month-by-month revenue analysis.
- **Automated Tasks:** VBA scripts to reduce manual effort and accelerate decision-making processes.

## Screenshots

### Time Frame Dashboard

![image](https://github.com/tushar11720/Excel-Interactive-Business-Intelligence-Dashboard/assets/132842128/70ee0bb6-567a-45a7-a443-de7d2a42c61b)

*Time Frame Dashboard: Analysis of profit trends over time, including month-over-month growth rates.*

### Store Dashboard

![image](https://github.com/tushar11720/Excel-Interactive-Business-Intelligence-Dashboard/assets/132842128/a7c31a8b-8e8e-4c35-a6eb-cba6e8aef958)

*Store Dashboard: Comparison of revenue vs. target for each store and month-by-month revenue analysis.*

### Profit Dashboard

![image](https://github.com/tushar11720/Excel-Interactive-Business-Intelligence-Dashboard/assets/132842128/aab97aaf-8d37-4fdd-83d2-a376e3bc2a44)
*Profit Dashboard: Insights into overall profitability, including product profitability and return rates.*

## VBA Code

The project includes a VBA module that provides additional functionality for the dashboard. Below is a snippet of the VBA code used to toggle the visibility of dashboard elements:

```vba
Sub ToggleVisibility()
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes("Group 9242")

    If shp.Visible = msoTrue Then
        shp.Visible = msoFalse
    Else
        shp.Visible = msoTrue
    End If
End Sub
```

## Conclusion

This interactive business intelligence dashboard leverages Excel's powerful data visualization and VBA scripting capabilities to provide valuable insights into business operations. By automating repetitive tasks and enhancing data exploration, the dashboard significantly improves operational efficiency and decision-making processes. It serves as a critical tool for data-driven decision-making and operational management, ultimately contributing to increased profitability and streamlined business operations.
