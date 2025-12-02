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
        this.addSpellPowerSourceBtn = get('add-spell-power-source-btn');
        this.empowerCheckbox = get('metamagic-empower');
        this.maximizeCheckbox = get('metamagic-maximize');
        this.intensifyCheckbox = get('metamagic-intensify');
        this.wellspringCheckbox = get('boost-wellspring');
        this.calculateBtn = get('calculate-spell-btn');

        // Output elements
        this.spellPowerSourcesContainer = get('spell-power-sources-container');
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
            this.wellspringCheckbox.addEventListener('change', recordChange);

            // Add listener for adding a new spell damage row
            this.addSpellDamageRowBtn.addEventListener('click', (e) => this.addSpellDamageRow(e));
            this.addSpellPowerSourceBtn.addEventListener('click', () => this.addSpellPowerSource());



            // Use event delegation for remove buttons within the spellDamageRowsContainer
            this.spellDamageRowsContainer.addEventListener('click', (e) => {
                if (e.target && e.target.classList.contains('remove-row-btn')) {
                    e.preventDefault();
                    const mainRowToRemove = e.target.closest('.spell-damage-source-row');
                    if (mainRowToRemove) {
                        // Remove all subsequent additional-scaling-row siblings
                        let nextSibling = mainRowToRemove.nextElementSibling;
                        while (nextSibling && nextSibling.classList.contains('additional-scaling-row')) {
                            const toRemove = nextSibling;
                            nextSibling = nextSibling.nextElementSibling;
                            toRemove.remove();
                        }
                        mainRowToRemove.remove(); // Remove the main row itself
                    }
                    this.calculateSpellDamage();
                    this.manager.saveState();
                } else if (e.target && e.target.classList.contains('add-scaling-input-btn')) {
                    e.preventDefault();
                    this._addAdditionalScalingInput(e.target.closest('.input-group-row'));
                }
            });

            // Use event delegation for remove buttons on spell power profiles
            this.spellPowerSourcesContainer.addEventListener('click', (e) => {
                if (e.target && e.target.classList.contains('remove-row-btn')) {
                    e.preventDefault();
                    const profileElement = e.target.closest('.spell-power-profile');
                    if (profileElement) {
                        this.removeSpellPowerSource(profileElement);
                    }
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

        // Find existing additional scaling rows that are siblings to the main row
        let existingScalingRows = 0;
        let nextSibling = rowElement.nextElementSibling;
        while (nextSibling && nextSibling.classList.contains('additional-scaling-row')) {
            existingScalingRows++;
            nextSibling = nextSibling.nextElementSibling;
        }

        const newScalingIndex = existingScalingRows; // The index for the new row (0-based)

        // Create label and input
        const spellPowerOptions = this._getSpellPowerProfiles().map(p => `<option value="${p.id}">SP ${p.id}</option>`).join('');
        const newAdditionalScalingId = existingScalingRows + 1; // A unique ID for the new elements
        const wrapper = document.createElement('div');
        wrapper.className = 'input-group-row additional-scaling-row';
        wrapper.innerHTML = `
            <label for="additional-scaling-base-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" class="short-label">SP ${newScalingIndex + 2} Base</label>
            <input type="text" id="additional-scaling-base-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Base damage for this component (does not scale with CL)">
            <span class="plus-symbol">+</span>
            <label for="additional-scaling-cl-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" class="short-label">per CL</label>
            <input type="text" id="additional-scaling-cl-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Bonus damage per caster level for this component">
            <select id="additional-scaling-sp-select-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" class="small-input" title="Select Spell Power source for this component">${spellPowerOptions}</select>            
            <button class="remove-scaling-input-btn small-btn" title="Remove this scaling input">&times;</button>
        `;

        // Insert the new row after the last existing scaling row, or after the main row if none exist.
        let insertionPoint = rowElement;
        if (existingScalingRows > 0) {
            insertionPoint = rowElement.nextElementSibling;
            for (let i = 1; i < existingScalingRows; i++) {
                insertionPoint = insertionPoint.nextElementSibling;
            }
        }
        insertionPoint.insertAdjacentElement('afterend', wrapper);

        // Set the default selection for the new dropdown
        const spSelect = wrapper.querySelector('select');
        if (spSelect) {
            // Default to the corresponding SP source if it exists, otherwise the last one.
            const targetProfileId = newScalingIndex + 2;
            const profiles = this._getSpellPowerProfiles();
            if (profiles.some(p => p.id === targetProfileId)) {
                spSelect.value = targetProfileId;
            } else if (profiles.length > 0) {
                spSelect.value = profiles[profiles.length - 1].id;
            }
        }

        // If 5 inputs are present, hide the add button
        if (existingScalingRows + 1 >= 5) {
            rowElement.querySelector('.add-scaling-input-btn').classList.add('hidden');
        }

        // Add event listener for the new remove button
        wrapper.querySelector('.remove-scaling-input-btn').addEventListener('click', (e) => {
            e.preventDefault();
            wrapper.remove();
            // After removing, ensure the main row's add button is visible again
            const addBtn = rowElement.querySelector('.add-scaling-input-btn');
            if (addBtn) {
                addBtn.classList.remove('hidden');
            }

            this.calculateSpellDamage();
            this.manager.saveState();
        });

        // Recalculate damage and save state
        this.calculateSpellDamage();
        this.manager.saveState();
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
        newRow.className = 'input-group-row spell-damage-source-row';
        newRow.innerHTML = `
            <input type="text" id="spell-name-${newRowId}${this.idSuffix}" value="Source ${newRowId}" title="Name of the spell component" placeholder="Spell Name" class="adaptive-text-input">
            <label for="spell-damage-${newRowId}${this.idSuffix}">Base Damage</label>
            <input type="text" id="spell-damage-${newRowId}${this.idSuffix}" value="0" title="The spell's base damage (e.g., 10d6+50)" class="adaptive-text-input">
            <span class="plus-symbol">+</span>
            <label for="spell-cl-scaling-${newRowId}${this.idSuffix}" class="short-label">per CL</label>
            <input type="text" id="spell-cl-scaling-${newRowId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Bonus damage dice per caster level (e.g., 1d6 per CL)">
            <label for="caster-level-${newRowId}${this.idSuffix}" class="short-label">CL</label>
            <input type="number" id="caster-level-${newRowId}${this.idSuffix}" value="20" class="small-input" title="Caster Level for this damage component">
            <label for="spell-hit-count-${newRowId}${this.idSuffix}" class="short-label">Hits</label>
            <input type="number" id="spell-hit-count-${newRowId}${this.idSuffix}" value="1" min="1" class="small-input" title="Number of times this spell component hits">
            <button class="add-scaling-input-btn small-btn" title="Add additional scaling input">+</button>
            <button class="remove-row-btn" title="Remove this damage source">&times;</button>
        `;
        this.spellDamageRowsContainer.appendChild(newRow);
        newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
    }

    _getInputs() {
        const spellPowerProfiles = this._getSpellPowerProfiles();
        const spellDamageSources = [];
        
        // Find only the main spell damage rows, not the additional scaling rows
        this.spellDamageRowsContainer.querySelectorAll('.input-group-row:not(.additional-scaling-row)').forEach((row, i) => {
            const spellNameInput = row.querySelector('input[id^="spell-name-"]');
            const baseDmgInput = row.querySelector('input[id^="spell-damage-"]');
            const clScalingInput = row.querySelector('input[id^="spell-cl-scaling-"]');
            const casterLevelInput = row.querySelector('input[id^="caster-level-"]');
            const hitCountInput = row.querySelector('input[id^="spell-hit-count-"]');

            const additionalScalings = [];
            let nextSibling = row.nextElementSibling;
            while (nextSibling && nextSibling.classList.contains('additional-scaling-row')) {
                const baseInput = nextSibling.querySelector('input[id^="additional-scaling-base-"]');
                const clInput = nextSibling.querySelector('input[id^="additional-scaling-cl-"]');
                const spSelect = nextSibling.querySelector('select[id^="additional-scaling-sp-select-"]');

                if (baseInput && clInput && spSelect) {
                    additionalScalings.push({
                        base: this.parseDiceNotation(baseInput.value),
                        clScaled: this.parseDiceNotation(clInput.value),
                        profileId: parseInt(spSelect.value, 10) || 1
                    });
                }
                nextSibling = nextSibling.nextElementSibling;
            }

            if (spellNameInput && baseDmgInput && clScalingInput && casterLevelInput && hitCountInput) {
                spellDamageSources.push({
                    name: spellNameInput.value || `Source ${i + 1}`,
                    base: this.parseDiceNotation(baseDmgInput.value),
                    clScaled: this.parseDiceNotation(clScalingInput.value),
                    casterLevel: parseInt(casterLevelInput.value) || 0,
                    hitCount: parseInt(hitCountInput.value) || 1,
                    additionalScalings: additionalScalings
                });
            }
        });

        const inputs = {
            spellDamageSources: spellDamageSources,
            spellPowerProfiles: spellPowerProfiles,

            isEmpowered: this.empowerCheckbox.checked,
            isMaximized: this.maximizeCheckbox.checked,
            isIntensified: this.intensifyCheckbox.checked,
            isWellspring: this.wellspringCheckbox.checked,
        };
        return inputs;
    }

    _getSpellPowerProfiles() {
        const profiles = [];
        this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile').forEach(profileEl => {
            const profileId = parseInt(profileEl.dataset.profileId, 10);
            const spellPower = parseInt(profileEl.querySelector(`#spell-power-${profileId}${this.idSuffix}`).value, 10) || 0;
            const critChance = (parseFloat(profileEl.querySelector(`#spell-crit-chance-${profileId}${this.idSuffix}`).value, 10) || 0) / 100;
            const critDamage = (parseFloat(profileEl.querySelector(`#spell-crit-damage-${profileId}${this.idSuffix}`).value, 10) || 0) / 100;

            profiles.push({
                id: profileId,
                spellPower: spellPower,
                critChance: critChance,
                critDamage: critDamage
            });
        });
        return profiles;
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

        const wellspringSpellPowerBonus = inputs.isWellspring ? 150 : 0;
        const wellspringCritDamageBonus = inputs.isWellspring ? 0.20 : 0;

        const totalSpellPowerBonus = metamagicSpellPower + wellspringSpellPowerBonus;
 
        // Update bonus displays for each spell power profile
        this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile').forEach(profileEl => {
            const profileId = profileEl.dataset.profileId;
            const spellPowerBonusSpan = profileEl.querySelector(`#metamagic-spell-power-bonus-${profileId}${this.idSuffix}`);
            const critDamageBonusSpan = profileEl.querySelector(`#boost-crit-damage-bonus-${profileId}${this.idSuffix}`);
            if (spellPowerBonusSpan) spellPowerBonusSpan.textContent = totalSpellPowerBonus;
            if (critDamageBonusSpan) critDamageBonusSpan.textContent = `${wellspringCritDamageBonus * 100}%`;
        });

        inputs.spellDamageSources.forEach(source => {
            let sourceTotalAverage = 0;
            const components = [];

            // --- Calculate Base + CL damage ---
            const baseDamageComponent = source.base + (source.clScaled * source.casterLevel);

            // Base damage always uses the first spell power source (SP 1)
            const baseProfile = inputs.spellPowerProfiles.find(p => p.id === 1) || inputs.spellPowerProfiles[0];
            if (baseProfile) {
                const totalSpellPower = baseProfile.spellPower + totalSpellPowerBonus;
                const spellPowerMultiplier = 1 + (totalSpellPower / 100);
                const critMultiplier = 2 + baseProfile.critDamage + wellspringCritDamageBonus;

                const averageHit = baseDamageComponent * spellPowerMultiplier;
                const averageCrit = averageHit * critMultiplier;
                const totalAverage = (averageHit * (1 - baseProfile.critChance)) + (averageCrit * baseProfile.critChance);

                sourceTotalAverage += totalAverage; // Keep track of the running total for the whole source
                components.push({
                    name: 'Base + CL',
                    averageHit: averageHit,
                    averageCrit: averageCrit
                });
            }

            // --- Calculate Additional Scaling damage ---
            source.additionalScalings.forEach((scaling, index) => {
                const profile = inputs.spellPowerProfiles.find(p => p.id === scaling.profileId) || inputs.spellPowerProfiles[0];
                if (profile) {
                    const totalSpellPower = profile.spellPower + totalSpellPowerBonus;
                    const spellPowerMultiplier = 1 + (totalSpellPower / 100);
                    const critMultiplier = 2 + profile.critDamage + wellspringCritDamageBonus;

                    // Combine base and CL-scaled damage for this component
                    const damageComponent = scaling.base + (scaling.clScaled * source.casterLevel);
                    const averageHit = damageComponent * spellPowerMultiplier;
                    const averageCrit = averageHit * critMultiplier;
                    const totalAverage = (averageHit * (1 - profile.critChance)) + (averageCrit * profile.critChance);

                    sourceTotalAverage += totalAverage; // Add to the running total
                    components.push({
                        name: `SP ${index + 2}`,
                        averageHit: averageHit,
                        averageCrit: averageCrit
                    });
                }
            });

            // Multiply the total average damage for this source by its hit count
            sourceTotalAverage *= source.hitCount;

            individualSpellDamages.push({
                name: source.name,
                components: components,
                totalAverage: sourceTotalAverage
            });
            totalBaseDamage += sourceTotalAverage;
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
        this.updateSpellPowerSelectors();
    }

    _updateSummaryUI() {
        const individualSpellDamageSummary = document.getElementById(`individual-spell-damage-summary${this.idSuffix}`);
        if (!individualSpellDamageSummary) return;

        individualSpellDamageSummary.innerHTML = '<h3>Individual Spell Damage</h3>'; // Clear previous results

        this.individualSpellDamages.forEach(spell => {
            const spellContainer = document.createElement('div');
            spellContainer.style.marginBottom = '0.5rem';
            spellContainer.innerHTML = `<strong>${spell.name} (Total Avg: ${spell.totalAverage.toFixed(2)})</strong>`;

            spell.components.forEach(component => {
                const p = document.createElement('p');
                p.style.marginLeft = '1rem';
                p.innerHTML = `<em>${component.name}:</em> Avg Hit: ${component.averageHit.toFixed(2)}, Avg Crit: ${component.averageCrit.toFixed(2)}`;
                spellContainer.appendChild(p);
            });
            individualSpellDamageSummary.appendChild(spellContainer);
        });
    }

    addSpellPowerSource(state = null) {
        const existingProfiles = this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile');
        let newProfileId;

        if (state && state.id) {
            newProfileId = state.id;
        } else {
            newProfileId = existingProfiles.length > 0 ? Math.max(...Array.from(existingProfiles).map(p => parseInt(p.dataset.profileId, 10))) + 1 : 1;
        }

        const newProfile = document.createElement('div');
        newProfile.className = 'stats-group spell-power-profile';
        newProfile.dataset.profileId = newProfileId;

        // Use state values if provided, otherwise use defaults
        const typeValue = state?.type || '';
        const powerValue = state?.spellPower || '500';
        const chanceValue = state?.critChance || '20';
        const damageValue = state?.critDamage || '100';

        newProfile.innerHTML = `
            <div class="profile-header">
                <label class="section-headline">Spell Power & Criticals ${newProfileId}</label>
                <input type="text" id="spell-power-type-${newProfileId}${this.idSuffix}" class="adaptive-text-input"
                    placeholder="Element Type" title="e.g., Fire, Acid, etc." value="${typeValue}">
            </div>
            <div class="input-group-row">
                <label for="spell-power-${newProfileId}${this.idSuffix}">Spell Power</label>
                <input type="number" id="spell-power-${newProfileId}${this.idSuffix}" value="${powerValue}" class="small-input" title="Your spell power for this profile.">
                <span class="plus-symbol">+</span>
                <span id="metamagic-spell-power-bonus-${newProfileId}${this.idSuffix}" class="read-only-bonus">0</span>
            </div>
            <div class="input-group-row">
                <label for="spell-crit-chance-${newProfileId}${this.idSuffix}">Crit Chance %</label>
                <input type="number" id="spell-crit-chance-${newProfileId}${this.idSuffix}" value="${chanceValue}" class="small-input" title="Your chance to critically hit with this profile.">
            </div>
            <div class="input-group-row">
                <label for="spell-crit-damage-${newProfileId}${this.idSuffix}">Crit Dmg Bonus %</label>
                <input type="number" id="spell-crit-damage-${newProfileId}${this.idSuffix}" value="${damageValue}" class="small-input" title="Your additional spell critical damage bonus. Base critical damage is +100% (x2 total).">
                <span class="plus-symbol">+</span>
                <span id="boost-crit-damage-bonus-${newProfileId}${this.idSuffix}" class="read-only-bonus">0</span>
            </div>
            <button class="remove-row-btn" title="Remove this profile">&times;</button>
        `;

        this.spellPowerSourcesContainer.appendChild(newProfile);

        // Only trigger recalculation and save if we are NOT in a bulk-loading state (i.e., state is null)
        if (state === null) {
            this.updateSpellPowerSelectors();
            this.calculateSpellDamage();
            this.manager.saveState();
        }
    }

    removeSpellPowerSource(profileElement) {
        if (this.spellPowerSourcesContainer.children.length <= 1) {
            alert("You cannot remove the last spell power source.");
            return;
        }
        profileElement.remove();
        this.updateSpellPowerSelectors();
        this.calculateSpellDamage();
        this.manager.saveState();
    }

    updateSpellPowerSelectors() {
        const profiles = this._getSpellPowerProfiles();
        const profileOptions = profiles.map(p => `<option value="${p.id}">SP ${p.id}</option>`).join('');

        document.querySelectorAll(`#calculator-set-${this.setId} select[id^="additional-scaling-sp-select-"]`).forEach(select => {
            const currentValue = select.value;
            select.innerHTML = profileOptions;

            // Try to preserve the selected value if it still exists
            if (profiles.some(p => p.id.toString() === currentValue)) {
                select.value = currentValue;
            } else {
                // If the selected profile was deleted, default to the first available one
                if (profiles.length > 0) {
                    select.value = profiles[0].id;
                }
            }
        });

        // Show/hide remove buttons
        const allProfiles = this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile');
        allProfiles.forEach(p => {
            p.querySelector('.remove-row-btn').classList.toggle('hidden', allProfiles.length <= 1);
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
            const hitCountInput = row.querySelector('input[id^="spell-hit-count-"]');

            const additionalScalingsState = [];
            let nextSibling = row.nextElementSibling;
            while (nextSibling && nextSibling.classList.contains('additional-scaling-row')) {
                const baseInput = nextSibling.querySelector('input[id^="additional-scaling-base-"]');
                const clInput = nextSibling.querySelector('input[id^="additional-scaling-cl-"]');
                const spSelect = nextSibling.querySelector('select[id^="additional-scaling-sp-select-"]');
                if (baseInput && clInput && spSelect) {
                    additionalScalingsState.push({
                        base: baseInput.value,
                        clScaled: clInput.value,
                        profileId: spSelect.value
                    });
                }
                nextSibling = nextSibling.nextElementSibling;
            }

            if (spellNameInput && baseDmgInput && clScalingInput && casterLevelInput && hitCountInput) {
                spellDamageSourcesState.push({
                    name: spellNameInput.value,
                    base: baseDmgInput.value,
                    clScaled: clScalingInput.value,
                    additionalScalings: additionalScalingsState,
                    casterLevel: casterLevelInput.value,
                    hitCount: hitCountInput.value
                });
            }
        });
        state.spellDamageSources = spellDamageSourcesState;

        // Store spell power profiles
        const spellPowerProfilesState = [];
        this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile').forEach(profileEl => {
            const profileId = parseInt(profileEl.dataset.profileId, 10);
            const spellPowerInput = profileEl.querySelector(`#spell-power-${profileId}${this.idSuffix}`);
            const critChanceInput = profileEl.querySelector(`#spell-crit-chance-${profileId}${this.idSuffix}`);
            const critDamageInput = profileEl.querySelector(`#spell-crit-damage-${profileId}${this.idSuffix}`);
            const typeInput = profileEl.querySelector(`#spell-power-type-${profileId}${this.idSuffix}`);

            if (spellPowerInput && critChanceInput && critDamageInput && typeInput) {
                spellPowerProfilesState.push({
                    id: profileId,
                    spellPower: spellPowerInput.value,
                    critChance: critChanceInput.value,
                    critDamage: critDamageInput.value,
                    type: typeInput.value
                });
            }
        });
        state.spellPowerProfiles = spellPowerProfilesState;

        return state;
    }

    setState(state) {
        if (!state) return;

        const allInputs = document.querySelectorAll(`#calculator-set-${this.setId} input, #calculator-set-${this.setId} select`);
        allInputs.forEach(input => {
            const key = input.id.replace(`-set${this.setId}`, '');
            // Skip dynamic inputs here, as they are handled below
            if (key.startsWith('spell-name-') || key.startsWith('spell-damage-') || key.startsWith('spell-cl-scaling-') || key.startsWith('additional-scaling-') || key.startsWith('caster-level-') || key.startsWith('spell-hit-count-') || key.startsWith('spell-power-type-') || key.startsWith('spell-power-') || key.startsWith('spell-crit-chance-') || key.startsWith('spell-crit-damage-') || key.startsWith('additional-scaling-sp-select-') || key.startsWith('boost-')) {
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

        // Handle spell power profiles
        const powerProfilesContainer = this.spellPowerSourcesContainer;
        powerProfilesContainer.innerHTML = ''; // Clear existing profiles

        if (state.spellPowerProfiles && state.spellPowerProfiles.length > 0) {
            // Recreate each profile directly from its state
            state.spellPowerProfiles.forEach(profileState => {
                this.addSpellPowerSource(profileState);
            });
        } else {
            // If no profiles in state, ensure one default one exists
            this.addSpellPowerSource(null);
        }

        // Handle spell damage rows
        // Clear all existing dynamic spell damage rows
        let existingRows = this.spellDamageRowsContainer.querySelectorAll('.input-group-row, .additional-scaling-row');
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
                    const hitCountInput = newRow.querySelector('input[id^="spell-hit-count-"]');

                    if (spellNameInput) spellNameInput.value = source.name;
                    if (baseDmgInput) baseDmgInput.value = source.base;
                    if (clScalingInput) clScalingInput.value = source.clScaled;
                    if (casterLevelInput) casterLevelInput.value = source.casterLevel;
                    if (hitCountInput) hitCountInput.value = source.hitCount;


                    // If additional scalings were saved, add the inputs and set their values
                    if (source.additionalScalings && source.additionalScalings.length > 0) {
                        source.additionalScalings.forEach((scalingState) => {
                            if (scalingState) {
                                this._addAdditionalScalingInput(newRow); // Add a new input field
                                const newScalingRow = newRow.nextElementSibling; // This is fragile, assumes it's added right after
                                const lastScalingRow = Array.from(this.spellDamageRowsContainer.querySelectorAll('.additional-scaling-row')).pop();

                                if (lastScalingRow) {
                                    const newBaseInput = lastScalingRow.querySelector('input[id^="additional-scaling-base-"]');
                                    const newClInput = lastScalingRow.querySelector('input[id^="additional-scaling-cl-"]');
                                    const newSpSelect = lastScalingRow.querySelector('select[id^="additional-scaling-sp-select-"]');
                                    if (newBaseInput) newBaseInput.value = scalingState.base;
                                    if (newClInput) newClInput.value = scalingState.clScaled;
                                    if (newSpSelect) newSpSelect.value = scalingState.profileId;
                                }
                            }
                        });
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
        this.updateSpellPowerSelectors();
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
