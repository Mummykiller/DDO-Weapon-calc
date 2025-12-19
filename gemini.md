# DDO Damage Calculator

This project implements a web-based Dungeons and Dragons Online (DDO) Damage Calculator. It allows users to input various combat statistics and character attributes to calculate and compare average damage per hit for different weapon and spell sets.

## Features

-   **Dynamic Calculator Sets:** Users can add multiple calculator sets to compare different weapon configurations, spell setups, or character builds side-by-side.
-   **Comprehensive Damage Calculation:** 
    -   **Weapon Calculation:** Takes into account weapon dice, bonus base damage, critical profile, miss/graze thresholds, Melee/Ranged Power, Sneak Attack, Imbues, and unscaled damage sources.
    -   **Spell Calculation:** A new calculator type for spells that accounts for base spell damage, Caster Level, Spell Power, Spell Critical Chance/Damage, and Metamagic feats (Empower, Maximize, Intensify). It also calculates damage reduction from a target's Magical Resistance Rating (MRR).
-   **Interactive UI:**
    -   Input fields for all relevant combat and spell stats.
    -   Real-time update of average damage results for both weapon and spell sets.
    -   Detailed breakdown of damage scaling.
    -   Table showing damage per d20 roll for weapon sets.
    -   Drag-and-drop functionality for reordering calculator sets.
    -   Tabbed navigation for switching between different calculator sets.
    -   Theming toggle (Light/Dark mode).
-   **Comparison Table:** A summary table displays the total average damage and components for all active calculator sets, highlighting the best performing set.
-   **State Management:** The calculator's state (all input values and active sets, including their type) is saved locally in the browser.
-   **Import/Export Functionality:** Users can import and export all calculator set data (both weapon and spell) in JSON format or as a human-readable summary.

## Technologies Used

-   **HTML5:** For the basic structure and content of the web page, including templates for both weapon and spell calculators.
-   **CSS3:** For styling the application, including responsive design, light/dark themes, and visual layout.
-   **JavaScript (ES6+):** For all interactive functionality, damage calculation logic, DOM manipulation, and state management. The application uses three main classes:
    -   `Calculator`: Handles the logic and UI for a single weapon damage calculator instance.
    -   `SpellCalculator`: Handles the logic and UI for a single spell damage calculator instance.
    -   `CalculatorManager`: Manages all `Calculator` and `SpellCalculator` instances, tab navigation, drag-and-drop, import/export, and overall application state.

## How it Works

The core logic revolves around parsing user inputs and applying the appropriate DDO damage formulas. 

For **weapon damage**, it calculates probabilities for different hit outcomes (miss, graze, normal hit, critical hit) and combines them with various damage components (base weapon damage, sneak attack, imbue, unscaled damage) and their respective scaling factors.

For **spell damage**, the logic calculates a base spell damage value and applies scaling from total Spell Power, which now includes bonuses from Metamagic feats. It then factors in spell critical chance and multiplier to determine the average damage before finally accounting for the target's MRR.

The UI is dynamically generated and updated using JavaScript, allowing for a flexible and interactive user experience. The use of local storage ensures persistence of user data across sessions.

## Project Artefacts

This repository contains `.toon` files, which are custom-formatted documents used for project planning and documentation.
-   `DDO_Damage_Calculator.toon`: Outlines the original architecture of the weapon damage calculator.
-   `DDO_Spell_Scaling.toon`: A knowledge base document detailing the core mechanics of spell damage scaling in DDO, based on web research.
-   `changelog.toon`: Documents the implementation of the spell calculator feature and other recent changes.