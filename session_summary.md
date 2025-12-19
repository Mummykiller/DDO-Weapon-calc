---
[Create Time]: 2025-11-30 18:57:40 UTC
[Title]: DDO Damage Calculator - Feature Enhancement and Bug Squashing
[Summary]:
- **Scaled Dice Feature Implementation:**
  - Added a "Scaled Dice Damage" section for weapon sets.
  - Implemented logic for dice to scale based on the number of Imbue Dice.
  - Added individual controls for each scaled dice row, including:
    - Enable/Disable checkbox.
    - Proc Chance percentage.
    - Melee/Ranged Power scaling percentage.
- **UI/UX Enhancements:**
  - Added a display in the input area to show how many extra dice are gained from imbues (e.g., "3 + 3").
  - Added a detailed calculation breakdown in the results area, showing how scaled dice damage is calculated, including subtotals and global multipliers (Multi-Strike, Reaper).
  - Fixed various CSS layout and text-wrapping issues for a cleaner presentation.
- **Critical Bug Fixes & Refactoring:**
  - **State Management:** Resolved a critical bug that caused a blank page on reload by correctly saving and loading dynamically added "Unscaled Damage" and "Scaled Dice" rows.
  - **Calculation Accuracy:**
    - Corrected multiple logical errors in the scaled dice calculation to ensure the number of added dice and their resulting damage were accurate.
    - Overhauled the main averaging function (`_calculateAverages`) to use a consistent, high-fidelity probability model for all damage types (Base, Sneak, Imbue, Unscaled, and Scaled), fixing a major inconsistency.
  - **Data Display:**
    - Fixed the "Damage per d20 Roll" table to correctly apply multipliers to scaled dice damage.
    - Updated the "Set Comparison" table to correctly display data for spell sets and include the "Avg Scaled" column.
  - **Code Health:**
    - Conducted a full code review to identify and fix remaining inconsistencies.
    - Resolved a bug in loading saved spell sets with additional scaling inputs.
    - Removed duplicated and erroneous code that was breaking script execution.
---