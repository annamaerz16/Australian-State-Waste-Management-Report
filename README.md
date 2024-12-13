# Australian-State-Waste-Management-Report

### Project Overview
Data Visualization  Australian State Waste Management Report using MS Excel. - Conditional Formatting and Macros

### Data Sources
https://www.environment.gov.au/system/files/resources/7381c1de-31d0-429b-912c-91a6dbc83af7/files/national-waste-report-2018.pdf

### Tools
Microsoft Excel

### KPI’s
- Core Waste Ton Production 2007-2017 (in 2 years steps)
- Resource Recovery Rates by Region 2007-2017 (in 2 years steps)
- Resource recovery and recycling rates 2016-2017 (in 2 years steps)


### Procedure
#### 1. Add Data
- Core Waste Ton PC 2007-2017 Table (in 2 years steps)
- Resource Recovery Rates by Region 2007-2017 Table (in 2 years steps)
- Resource recovery and recycling rates 2016-2017 Table (in 2 years steps)
- Drop-Down Menu with all the States in Australia
- Map of Australia

#### 2. Conditional Formatting
##### Core Waste Ton Table
- value-based conditional formatting: Bottom 5 Values
- value-based conditional formatting: Map highlights state in color if the state is selected in cell K1 (new rule--> formula, Format painter for all states)
- value-based conditional formatting: highlight row of the state that is selected in K1 (new rule --> formula, change order of rules to still see the highlighted bottom 5 values)

##### Resource Recovery Rates by Region Table
- trend-based conditional formatting: Data Bars
- trend-based conditional formatting:Iconsets (Edit rule: green: >0, yellow:=0, red:<0, show icons only) (shows increase/decrease growth rate) 
##### Resource Recovery and Recycling Rates Table
- Landfill: trend-based conditional formatting: colorscale (red-yellow-green-scale)
- value-based conditional formatting: Recycling Rates below 48%

#### 3. Macros
##### Core Waste Ton Table
- Link Map to Macro, so by clicking the state in the map, row values geht highlighted (record macro, add macros for all states (VBA), assign macros to map)



