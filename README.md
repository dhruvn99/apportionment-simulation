# Apportionment Simulation

Replication code for *Quantifying Representational Deviation in 
Divisor-Based Apportionment Methods* (Nimmagadda, 2026).

Senior Honors Thesis, Department of Mathematics, University of Connecticut.

## Overview

This repository contains the complete Python simulation used to evaluate 
six congressional apportionment methods — Jefferson, Adams, Webster, 
Hill-Huntington, Dean, and Hamilton — across historical U.S. Census data 
(1990–2020) and 5,000 Monte Carlo draws per synthetic scenario.

## Files

- `apportionment_simulation.py` — main simulation code
- `census_data.xlsx` — U.S. Census Bureau apportionment populations, 1990–2020
- `requirements.txt` — Python dependencies

## Requirements

Python 3.12 or later. Install dependencies with:

    pip install -r requirements.txt

## Usage

    python apportionment_simulation.py

Output files (CSV results and figures) are saved to the directory 
specified in `OUTPUT_DIR` at the top of the script.

## Reproducibility

All results are reproducible with random seed 42, set at the top of the script.
