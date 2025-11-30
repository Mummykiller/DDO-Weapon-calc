class SpellCalculator {
    constructor(setId, manager, name) {
        this.setId = setId;
        this.manager = manager;
        this.name = name;
        this.idSuffix = `-set${setId}`;

        // Properties for spell calculations
        this.totalAverageDamage = 0;
        this.averageBaseHit = 0;
        this.averageCritHit = 0;
        this.finalAverageDamage = 0;


        this.getElements();
        this.addEventListeners();

        // Create a hidden span for measuring text width
        this._measurementSpan = document.createElement('span');
        this._measurementSpan.style.position = 'absolute';
        this._measurementSpan.style.visibility = 'hidden';
        this._measurementSpan.style.whiteSpace = 'nowrap';
        document.body.appendChild(this._measurementSpan);

        this._initializeAdaptiveInputs();

        // Initial calculation
        this.calculateSpellDamage();
    }

    getElements() {
        const get = (baseId) => document.getElementById(baseId + this.idSuffix);

        // Input elements
        this.spellDamageRowsContainer = get('spell-damage-rows-container');
        this.addSpellDamageRowBtn = get('add-spell-damage-row-btn');

        this.spellPowerInput = get('spell-power');
        this.spellCritChanceInput = get('spell-crit-chance');
        this.spellCritDamageInput = get('spell-crit-damage');

        this.empowerCheckbox = get('metamagic-empower');
        this.maximizeCheckbox = get('metamagic-maximize');
        this.intensifyCheckbox = get('metamagic-intensify');
        this.calculateBtn = get('calculate-spell-btn');

        // Output elements
        this.metamagicSpellPowerBonusSpan = get('metamagic-spell-power-bonus');
        this.avgSpellDamageSpan = get('avg-spell-damage');
        this.avgSpellCritDamageSpan = get('avg-spell-crit-damage');
        this.totalAvgSpellDamageSpan = get('total-avg-spell-damage');

        this.finalSpellDamageSpan = get('final-spell-damage');
    }

    addEventListeners() {
        this.calculateBtn.addEventListener('click', () => this.calculateSpellDamage());

        const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
        if (calculatorElement) {
            const recordChange = (e) => {
                this.calculateSpellDamage();
                this.manager.saveState();
            };
            calculatorElement.addEventListener('input', recordChange);
            calculatorElement.addEventListener('change', recordChange);

            // Add direct listeners for metamagic checkboxes
            this.empowerCheckbox.addEventListener('change', recordChange);
            this.maximizeCheckbox.addEventListener('change', recordChange);
            this.intensifyCheckbox.addEventListener('change', recordChange);

            // Add listener for adding a new spell damage row
            this.addSpellDamageRowBtn.addEventListener('click', (e) => this.addSpellDamageRow(e));



            // Use event delegation for remove buttons within the spellDamageRowsContainer
            this.spellDamageRowsContainer.addEventListener('click', (e) => {
                if (e.target && e.target.classList.contains('remove-row-btn')) {
                    e.preventDefault();
                    e.target.closest('.input-group-row').remove();
                    this.calculateSpellDamage(); // Recalculate after removing a row
                    this.manager.saveState();
                } else if (e.target && e.target.classList.contains('add-scaling-input-btn')) {
                    e.preventDefault();
                    this._addAdditionalScalingInput(e.target.closest('.input-group-row'));
                } else if (e.target && e.target.classList.contains('remove-scaling-input-btn')) {
                    e.preventDefault();
                    this._removeAdditionalScalingInput(e.target.closest('.input-group-row'));
                }
            });
        }
    }

    removeEventListeners() {
        // Event listeners are on the container, which gets removed, 
        // so no specific removal is needed here with the current event delegation model.
    }

    _resizeInput(inputElement) {
        // Apply relevant styles from the input to the measurement span
        const computedStyle = window.getComputedStyle(inputElement);
        this._measurementSpan.style.fontFamily = computedStyle.fontFamily;
        this._measurementSpan.style.fontSize = computedStyle.fontSize;
        this._measurementSpan.style.fontWeight = computedStyle.fontWeight;
        this._measurementSpan.style.letterSpacing = computedStyle.letterSpacing;
        this._measurementSpan.style.textTransform = computedStyle.textTransform;

        const paddingLeft = parseFloat(computedStyle.paddingLeft);
        const paddingRight = parseFloat(computedStyle.paddingRight);
        const borderWidthLeft = parseFloat(computedStyle.borderLeftWidth);
        const borderWidthRight = parseFloat(computedStyle.borderRightWidth);

        this._measurementSpan.textContent = inputElement.value || inputElement.placeholder || '';

        // Calculate the desired width including padding and border and a buffer
        let desiredWidth = this._measurementSpan.offsetWidth + paddingLeft + paddingRight + borderWidthLeft + borderWidthRight + 4;

        const minWidth = parseFloat(computedStyle.minWidth) || 50;

        inputElement.style.width = `${Math.max(minWidth, desiredWidth)}px`;
    }


    _initializeAdaptiveInputs() {
        const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
        if (!calculatorElement) return;

        // Use event delegation
        calculatorElement.addEventListener('input', (e) => {
            if (e.target.classList.contains('adaptive-text-input')) {
                this._resizeInput(e.target);
            }
        });

        // Run initial resize on all existing inputs
        this.resizeAllAdaptiveInputs();
    }

    resizeAllAdaptiveInputs() {
        const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
        if (!calculatorElement) return;
        calculatorElement.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
    }


    _addAdditionalScalingInput(rowElement) {
        const rowIdMatch = rowElement.querySelector('input[id^="spell-name-"]').id.match(/spell-name-(\d+)/);
        if (!rowIdMatch) return;
        const spellRowId = rowIdMatch[1];

        const additionalScalingInputs = rowElement.querySelectorAll('input[id^="additional-scaling-"]');
        if (additionalScalingInputs.length >= 5) { // Limit to 5
            rowElement.querySelector('.add-scaling-input-btn').classList.add('hidden');
            return;
        }

        const addScalingButton = rowElement.querySelector('.add-scaling-input-btn');
        const removeScalingButton = rowElement.querySelector('.remove-scaling-input-btn');

        // Create label and input
        const newAdditionalScalingId = additionalScalingInputs.length + 1;
        const label = document.createElement('label');
        label.setAttribute('for', `additional-scaling-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}`);
        label.classList.add('short-label');
        label.textContent = `Add Scale ${newAdditionalScalingId}`;

        const input = document.createElement('input');
        input.setAttribute('type', 'text');
        input.setAttribute('id', `additional-scaling-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}`);
        input.setAttribute('value', '0');
        input.classList.add('small-input', 'adaptive-text-input');
        input.setAttribute('title', 'Additional flat damage bonus applied after Spell Power scaling');

        // Insert elements before the addScalingButton
        rowElement.insertBefore(label, addScalingButton);
        rowElement.insertBefore(input, addScalingButton);

        // Show the remove button
        removeScalingButton.classList.remove('hidden');

        // If 5 inputs are present, hide the add button
        if (additionalScalingInputs.length + 1 >= 5) {
            addScalingButton.classList.add('hidden');
        }

        // Recalculate damage and save state
        this.calculateSpellDamage();
        this.manager.saveState();
        this._resizeInput(input); // Resize the new input
    }

    _removeAdditionalScalingInput(rowElement) {
        const additionalScalingInputs = rowElement.querySelectorAll('input[id^="additional-scaling-"]');
        if (additionalScalingInputs.length === 0) {
            return; // Nothing to remove
        }

        const lastInput = additionalScalingInputs[additionalScalingInputs.length - 1];
        const lastLabel = rowElement.querySelector(`label[for="${lastInput.id}"]`);

        if (lastInput) lastInput.remove();
        if (lastLabel) lastLabel.remove();

        const addScalingButton = rowElement.querySelector('.add-scaling-input-btn');
        const removeScalingButton = rowElement.querySelector('.remove-scaling-input-btn');

        // If no more inputs, hide the remove button
        if (additionalScalingInputs.length - 1 === 0) {
            removeScalingButton.classList.add('hidden');
        }

        // Always show the add button if below limit
        if (additionalScalingInputs.length - 1 < 5) {
            addScalingButton.classList.remove('hidden');
        }

        this.calculateSpellDamage();
        this.manager.saveState();
    }


    addSpellDamageRow(e) {
        e.preventDefault();

        let maxRowId = 0;
        this.spellDamageRowsContainer.querySelectorAll('.input-group-row').forEach(row => {
            const firstInput = row.querySelector('input[id^="spell-name-"]');
            if (firstInput) {
                const idNum = parseInt(firstInput.id.match(/spell-name-(\d+)/)[1], 10);
                if (idNum > maxRowId) {
                    maxRowId = idNum;
                }
            }
        });
        const newRowId = maxRowId + 1;

        const newRow = document.createElement('div');
        newRow.className = 'input-group-row';
        newRow.innerHTML = `
            <input type="text" id="spell-name-${newRowId}${this.idSuffix}" value="Source ${newRowId}" title="Name of the spell component" placeholder="Spell Name" class="adaptive-text-input">
            <label for="spell-damage-${newRowId}${this.idSuffix}">Base Damage</label>
            <input type="text" id="spell-damage-${newRowId}${this.idSuffix}" value="0" title="The spell's base damage (e.g., 10d6+50)" class="adaptive-text-input">
            <span class="plus-symbol">+</span>
            <label for="spell-cl-scaling-${newRowId}${this.idSuffix}" class="short-label">per CL</label>
            <input type="text" id="spell-cl-scaling-${newRowId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Bonus damage dice per caster level (e.g., 1d6 per CL)">
            <label for="caster-level-${newRowId}${this.idSuffix}" class="short-label">CL</label>
            <input type="number" id="caster-level-${newRowId}${this.idSuffix}" value="20" class="small-input" title="Caster Level for this damage component">
            <button class="add-scaling-input-btn small-btn" title="Add additional scaling input">+</button>
            <button class="remove-scaling-input-btn small-btn hidden" title="Remove last additional scaling input">-</button>
            <button class="remove-row-btn" title="Remove this damage source">&times;</button>
        `;
        this.spellDamageRowsContainer.appendChild(newRow);
        newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
    }

    _getInputs() {
        const spellDamageSources = [];
        this.spellDamageRowsContainer.querySelectorAll('.input-group-row').forEach((row, i) => {
            const spellNameInput = row.querySelector('input[id^="spell-name-"]');
            const baseDmgInput = row.querySelector('input[id^="spell-damage-"]');
            const clScalingInput = row.querySelector('input[id^="spell-cl-scaling-"]');
            const casterLevelInput = row.querySelector('input[id^="caster-level-"]');

            // Collect all additional scaling inputs for this row
            let totalAdditionalScaling = 0;
            row.querySelectorAll('input[id^="additional-scaling-"]').forEach(input => {
                totalAdditionalScaling += this.parseDiceNotation(input.value);
            });

            if (spellNameInput && baseDmgInput && clScalingInput && casterLevelInput) {
                spellDamageSources.push({
                    name: spellNameInput.value || `Source ${i + 1}`,
                    base: this.parseDiceNotation(baseDmgInput.value),
                    clScaled: this.parseDiceNotation(clScalingInput.value),
                    additionalScaling: totalAdditionalScaling, // Store the sum
                    casterLevel: parseInt(casterLevelInput.value) || 0
                });
            }
        });

        const inputs = {
            spellDamageSources: spellDamageSources,
            spellPower: parseInt(this.spellPowerInput.value) || 0,
            critChance: (parseFloat(this.spellCritChanceInput.value) || 0) / 100,
            critDamage: (parseFloat(this.spellCritDamageInput.value) || 0) / 100,

            isEmpowered: this.empowerCheckbox.checked,
            isMaximized: this.maximizeCheckbox.checked,
            isIntensified: this.intensifyCheckbox.checked,
        };
        return inputs;
    }

    calculateSpellDamage() {
        let totalBaseDamage = 0;
        const individualSpellDamages = []; // To store results for each spell
        const inputs = this._getInputs();

        let metamagicSpellPower = 0;
        if (inputs.isIntensified) {
            metamagicSpellPower += 75;
        }
        if (inputs.isEmpowered) {
            metamagicSpellPower += 75;
        }
        if (inputs.isMaximized) {
            metamagicSpellPower += 150;
        }

        this.metamagicSpellPowerBonusSpan.textContent = metamagicSpellPower;
        const totalSpellPower = inputs.spellPower + metamagicSpellPower;
        const spellPowerMultiplier = 1 + (totalSpellPower / 100);
        const critMultiplier = 2 + inputs.critDamage;

        inputs.spellDamageSources.forEach(source => {
            const spellBaseDamage = source.base + (source.clScaled * source.casterLevel);
            const damageBeforeSpellPower = spellBaseDamage + source.additionalScaling; // Add additional scaling here
            const averageHit = damageBeforeSpellPower * spellPowerMultiplier;
            const averageCrit = averageHit * critMultiplier;
            const totalAverage = (averageHit * (1 - inputs.critChance)) + (averageCrit * inputs.critChance);

            individualSpellDamages.push({
                name: source.name,
                averageHit: averageHit,
                averageCrit: averageCrit,
                totalAverage: totalAverage
            });
            totalBaseDamage += totalAverage; // Summing up the individual total averages for the grand total
        });

        // This is the aggregated average BEFORE MRR for all spells combined


        // Store for comparison table - this is the final damage for ALL spells combined
        this.totalAverageDamage = totalBaseDamage;
        this.individualSpellDamages = individualSpellDamages; // Store individual results

        // Update UI
        this.avgSpellDamageSpan.textContent = totalBaseDamage.toFixed(2); // This now represents total pre-MRR average
        this.totalAvgSpellDamageSpan.textContent = totalBaseDamage.toFixed(2);
        this.finalSpellDamageSpan.textContent = totalBaseDamage.toFixed(2);

        this._updateSummaryUI(); // Call the new UI update method
    }

    _updateSummaryUI() {
        const individualSpellDamageSummary = document.getElementById(`individual-spell-damage-summary${this.idSuffix}`);
        if (!individualSpellDamageSummary) return;

        individualSpellDamageSummary.innerHTML = '<h3>Individual Spell Damage</h3>'; // Clear previous results

        this.individualSpellDamages.forEach(spell => {
            const p = document.createElement('p');
            p.innerHTML = `<strong>${spell.name}:</strong> Avg Dmg (pre-crit): ${spell.averageHit.toFixed(2)}, Avg Crit: ${spell.averageCrit.toFixed(2)}`;
            individualSpellDamageSummary.appendChild(p);
        });
    }

    getState() {
        const state = {};
        const allInputs = document.querySelectorAll(`#calculator-set-${this.setId} input, #calculator-set-${this.setId} select`);
        allInputs.forEach(input => {
            const key = input.id.replace(`-set${this.setId}`, '');
            if (input.type === 'checkbox') {
                state[key] = input.checked;
            } else {
                state[key] = input.value;
            }
        });

        // Store spell damage rows separately
        const spellDamageSourcesState = [];
        this.spellDamageRowsContainer.querySelectorAll('.input-group-row').forEach((row) => {
            const spellNameInput = row.querySelector('input[id^="spell-name-"]');
            const baseDmgInput = row.querySelector('input[id^="spell-damage-"]');
            const clScalingInput = row.querySelector('input[id^="spell-cl-scaling-"]');
            const casterLevelInput = row.querySelector('input[id^="caster-level-"]');

            const additionalScalings = [];
            row.querySelectorAll('input[id^="additional-scaling-"]').forEach(input => {
                additionalScalings.push(input.value);
            });

            if (spellNameInput && baseDmgInput && clScalingInput && casterLevelInput) {
                spellDamageSourcesState.push({
                    name: spellNameInput.value,
                    base: baseDmgInput.value,
                    clScaled: clScalingInput.value,
                    additionalScalings: additionalScalings, // Store as an array
                    casterLevel: casterLevelInput.value
                });
            }
        });
        state.spellDamageSources = spellDamageSourcesState;

        return state;
    }

    setState(state) {
        if (!state) return;

        const allInputs = document.querySelectorAll(`#calculator-set-${this.setId} input, #calculator-set-${this.setId} select`);
        allInputs.forEach(input => {
            const key = input.id.replace(`-set${this.setId}`, '');
            // Skip dynamic spell damage source inputs here, as they are handled below
            if (key.startsWith('spell-name-') || key.startsWith('spell-damage-') || key.startsWith('spell-cl-scaling-') || key.startsWith('additional-scaling-') || key.startsWith('caster-level-')) {
                return;
            }

            if (state.hasOwnProperty(key)) {
                if (input.type === 'checkbox') {
                    input.checked = state[key];
                } else {
                    input.value = state[key];
                }
            }
        });

        // Handle spell damage rows
        // Clear all existing dynamic spell damage rows
        let existingRows = this.spellDamageRowsContainer.querySelectorAll('.input-group-row');
        existingRows.forEach(row => row.remove());

        if (state.spellDamageSources && state.spellDamageSources.length > 0) {
            state.spellDamageSources.forEach((source, index) => {
                // Call addSpellDamageRow to create the div with the correct dynamic IDs
                this.addSpellDamageRow(new Event('dummy')); // Pass a dummy event

                // Get the newly added row (always the last child)
                const newRow = this.spellDamageRowsContainer.lastElementChild;

                // Populate the inputs within this new row
                if (newRow) {
                    const spellNameInput = newRow.querySelector('input[id^="spell-name-"]');
                    const baseDmgInput = newRow.querySelector('input[id^="spell-damage-"]');
                    const clScalingInput = newRow.querySelector('input[id^="spell-cl-scaling-"]');
                    const casterLevelInput = newRow.querySelector('input[id^="caster-level-"]');

                    if (spellNameInput) spellNameInput.value = source.name;
                    if (baseDmgInput) baseDmgInput.value = source.base;
                    if (clScalingInput) clScalingInput.value = source.clScaled;
                    if (casterLevelInput) casterLevelInput.value = source.casterLevel;

                    // If additional scalings were saved, add the inputs and set their values
                    if (source.additionalScalings && source.additionalScalings.length > 0) {
                        source.additionalScalings.forEach((scalingValue, scalingIndex) => {
                            if (scalingValue) {
                                this._addAdditionalScalingInput(newRow); // Add a new input field
                                const allScalingInputs = newRow.querySelectorAll('input[id^="additional-scaling-"]');
                                const newScalingInput = allScalingInputs[scalingIndex]; // Get the specific input we just added
                                if (newScalingInput) newScalingInput.value = scalingValue;
                            }
                        });
                    }

                    // Adjust button visibility after all additional scalings are loaded
                    const currentAdditionalInputs = newRow.querySelectorAll('input[id^="additional-scaling-"]');
                    const addScalingButton = newRow.querySelector('.add-scaling-input-btn');
                    const removeScalingButton = newRow.querySelector('.remove-scaling-input-btn');

                    if (currentAdditionalInputs.length > 0) {
                        removeScalingButton.classList.remove('hidden');
                    } else {
                        removeScalingButton.classList.add('hidden');
                    }

                    if (currentAdditionalInputs.length >= 5) {
                        addScalingButton.classList.add('hidden');
                    } else {
                        addScalingButton.classList.remove('hidden');
                    }
                }
            });
        } else {
            // If no saved sources, ensure at least one default empty row if needed
            // This case should be handled by the initial template, but to be safe:
            if (this.spellDamageRowsContainer.children.length === 0) {
                this.addSpellDamageRow(new Event('dummy'));
            }
        }
        this.resizeAllAdaptiveInputs();
        this.calculateSpellDamage();
    }

    parseDiceNotation(diceString) {
        // Ensure we have a string, trim whitespace
        let cleanString = (diceString || '').trim();

        if (!cleanString) {
            return 0;
        }

        // First, find and replace all range notations (e.g., "100-300") with their average value.
        // This prevents the '-' in a range from being treated as subtraction.
        cleanString = cleanString.replace(/(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)/g, (match, minStr, maxStr) => {
            const min = parseFloat(minStr);
            const max = parseFloat(maxStr);
            return ((min + max) / 2).toString();
        });
        // Standardize operators: replace all '-' with '+-' to make splitting easier
        cleanString = cleanString.replace(/\s/g, '');
        // Handle negative numbers at the start of the string
        if (cleanString.startsWith('-')) {
            cleanString = cleanString.substring(1).replace(/-/g, '+-');
            cleanString = '-' + cleanString;
        } else {
            cleanString = cleanString.replace(/-/g, '+-');
        }

        // Split the string by the '+' operator to get all the terms
        const terms = cleanString.split('+');

        let totalAverage = 0;

        for (const term of terms) {
            if (!term) continue; // Skip empty terms that can result from " -5"

            // Check if the term is a die roll (e.g., "2d6")
            if (term.toLowerCase().includes('d')) {
                const parts = term.toLowerCase().split('d');
                if (parts.length !== 2) continue; // Invalid format, skip

                let numDice;
                if (parts[0] === '-') {
                    numDice = -1;
                } else if (parts[0] === '') {
                    numDice = 1;
                } else {
                    numDice = parseInt(parts[0], 10);
                }

                if (isNaN(numDice)) numDice = 1;

                const numSides = parseFloat(parts[1]);

                if (isNaN(numSides) || numSides <= 0) continue; // Invalid sides, skip

                // Average of one die is (sides + 1) / 2. Multiply by the number of dice.
                totalAverage += numDice * (numSides + 1) / 2;
            } else {
                // If not a die roll, it's a flat number (e.g., "5" or "-2")
                totalAverage += parseFloat(term) || 0;
            }
        }
        return totalAverage;
    }

    getTabName() {
        const tab = document.querySelector(`.nav-tab[data-set="${this.setId}"] .tab-name`);
        return tab ? tab.textContent : `Set ${this.setId}`;
    }


    setTabName(name) {
        const tab = document.querySelector(`.nav-tab[data-set="${this.setId}"] .tab-name`);
        if (tab) {
            tab.textContent = name;
        }
    }
}
