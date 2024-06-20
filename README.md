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

![Customer Analysis](path_to_customer_analysis_screenshot.png)
*Customer Analysis: Insights into profit by gender and age group.*

![Profitability Trends](path_to_profitability_trends_screenshot.png)
*Profitability Trends: Month-over-month growth rates.*

![Weekday Profitability](path_to_weekday_profitability_screenshot.png)
*Weekday Profitability: Most profitable days of the week.*

![Product Analysis](path_to_product_analysis_screenshot.png)
*Product Analysis: Top-selling and most profitable products.*

![Store Performance](path_to_store_performance_screenshot.png)
*Store Performance: Revenue vs. target comparison for each store.*

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
