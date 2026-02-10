# Investment Decision Support Model (Excel + VBA)

This project is an Excel-based investment decision support model that combines
deterministic financial modeling with Monte Carlo simulation and scenario analysis.
The model is designed to support data-driven investment decisions under uncertainty.

## Key Features
- Deterministic NPV calculation based on core financial assumptions
- Monte Carlo simulation to model uncertainty and outcome distributions
- Scenario analysis (Pessimistic / Realistic / Optimistic)
- Automated histogram visualization with downside/upside color coding
- Printable report view with automatic chart export
- VBA-driven workflow optimized for performance and usability

## Model Structure
The workbook is organized into clearly separated sheets:

- **Data & Calculations**  
  Core input parameters and main financial assumptions.

- **Calculations** (Hidden) 
  Central calculation logic. The Monte Carlo simulation reads the simulated output
  from a single result cell to ensure consistency across analyses.

- **MC_sim**  
  Simulation control panel and visualization (histogram / Chart 2).

- **MC_data**  
  Raw Monte Carlo simulation output (one result per simulation).

- **MC_hist**  
  Histogram bins, frequencies, and distribution calculations.

- **Report (print)**  
  Final report view designed for presentation and PDF export.  
  The histogram is automatically inserted here as a static image.

## Scenario Analysis
The model includes a predefined scenario framework:
- **Pessimistic**
- **Realistic**
- **Optimistic**

Scenario selection dynamically updates model inputs, allowing both
deterministic evaluation and Monte Carlo analysis under consistent assumptions.

## Monte Carlo Simulation
- Simulation count is user-defined
- The model temporarily switches to manual calculation mode for performance
- After simulation, calculation mode is fully restored to avoid stale values
- Histogram bars are automatically colored to distinguish negative and positive outcomes

## How to Use
1. Open `Investment_Decision_Support_Model.xlsm`
2. Enable macros
3. Select a scenario (Pessimistic / Realistic / Optimistic)
4. Set the number of simulations
5. Run the Monte Carlo macro
6. Review results in **Report (print)**

## Example Output
Screenshots of the model output and report view can be found in the `screenshots/` folder.

## Notes & Limitations
- The model is intended for decision support and educational / portfolio purposes

## Skills Demonstrated
- Financial modeling and valuation logic
- Risk and uncertainty analysis (Monte Carlo simulation)
- Scenario-based decision analysis
- Advanced Excel and VBA development
- Clear separation of logic, data, and presentation layers
