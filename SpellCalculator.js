import { parseDiceNotation } from './utils.js';
import { BaseCalculator } from './BaseCalculator.js';

const defaultState = {
    'metamagic-empower': true,
    'metamagic-maximize': true,
    'metamagic-intensify': true,
    'boost-wellspring': false,
    'boost-night-horrors': false,
    // Default spell damage row
    'spell-name-1': 'Spell 1',
    'spell-damage-1': '10d6',
    'spell-cl-scaling-1': '1d6',
    'caster-level-1': 20,
    'spell-hit-count-1': 1,
    // Default spell power profile
    'spell-power-type-1': '',
    'spell-power-1': 1000,
    'spell-crit-chance-1': 75,
    'spell-crit-damage-1': 205,
};

export class SpellCalculator extends BaseCalculator {
    constructor(setId, manager, name) {
        super(setId, manager, name);

        // In-memory state object
        this.state = JSON.parse(JSON.stringify(defaultState));

        // Properties for spell calculations
        this.averageBaseHit = 0;
        this.averageCritHit = 0;
        this.finalAverageDamage = 0;


        this.getElements();
        this.addEventListeners();

        this._initializeAdaptiveInputs();

        // Initial calculation
        this.calculateSpellDamage();
    }

    getElements() {
        const get = (elementName) => this.container.querySelector(`[data-element="${elementName}"]`);

        // Input elements
        this.spellDamageRowsContainer = get('spell-damage-rows-container');
        this.addSpellDamageRowBtn = get('add-spell-damage-row-btn');
        this.addSpellPowerSourceBtn = get('add-spell-power-source-btn');
        this.empowerCheckbox = get('metamagic-empower');
        this.maximizeCheckbox = get('metamagic-maximize');
        this.intensifyCheckbox = get('metamagic-intensify');
        this.wellspringCheckbox = get('boost-wellspring');
        this.nightHorrorsCheckbox = get('boost-night-horrors');
        this.calculateBtn = get('calculate-spell-btn');

        // Output elements
        this.spellPowerSourcesContainer = get('spell-power-sources-container');
        this.avgSpellDamageSpan = get('avg-spell-damage');
        this.avgSpellCritDamageSpan = get('avg-spell-crit-damage');
        this.totalAvgSpellDamageSpan = get('total-avg-spell-damage');
        this.individualSpellDamageSummary = get('individual-spell-damage-summary');
        this.finalSpellDamageSpan = get('final-spell-damage');
    }

    addEventListeners() {
        super.addEventListeners();
        this.calculateBtn.addEventListener('click', () => this.calculateSpellDamage());

        const boundHandler = this.handleInputChange.bind(this);
        this.container.addEventListener('input', boundHandler);
        this.container.addEventListener('change', boundHandler);


        if (this.container) {
            // Add listener for adding a new spell damage row
            this.addSpellDamageRowBtn.addEventListener('click', (e) => this.addSpellDamageRow(e));
            this.addSpellPowerSourceBtn.addEventListener('click', () => this.addSpellPowerSource());



            // Use event delegation for remove buttons within the spellDamageRowsContainer
            this.spellDamageRowsContainer.addEventListener('click', (e) => {
                const removeBtn = e.target.closest('.remove-row-btn');
                if (removeBtn) {
                    const groupToRemove = e.target.closest('.spell-source-group');
                    if (groupToRemove) {
                        e.preventDefault();
                        this.removeSpellDamageRow(groupToRemove);
                    }
                } else if (e.target.closest('.duplicate-row-btn')) {
                    e.preventDefault();
                    this.duplicateSpellDamageRow(e.target.closest('.spell-source-group'));
                } else if (e.target.closest('.add-scaling-input-btn')) {
                    e.preventDefault();
                    this._addAdditionalScalingInput(e.target.closest('.spell-source-group'));
                }
            });

            // Use event delegation for remove buttons on spell power profiles
            this.spellPowerSourcesContainer.addEventListener('click', (e) => {
                const removeBtn = e.target.closest('.remove-row-btn');
                if (removeBtn) {
                    e.preventDefault();
                    const profileElement = e.target.closest('.spell-power-profile');
                    if (profileElement) {
                        this.removeSpellPowerSource(profileElement);
                    }
                }
            });

            // Add event listener for spell power type input changes using event delegation
            this.spellPowerSourcesContainer.addEventListener('input', (e) => {
                if (e.target.matches('[data-element^="spell-power-type-"]')) {
                    this.handleInputChange(e);
                }
            });
        }
    }

    _addAdditionalScalingInput(sourceGroup) {
        const mainRow = sourceGroup.querySelector('.spell-damage-source-row');
        const rowIdMatch = mainRow.querySelector('input[data-element^="spell-name-"]').dataset.element.match(/(\d+)$/);
        if (!rowIdMatch) return;
        const spellRowId = rowIdMatch[1];

        const scalingsContainer = sourceGroup.querySelector('.additional-scalings-container');
        if (!scalingsContainer) return;

        const existingScalingRows = scalingsContainer.children.length;
        const newScalingIndex = existingScalingRows; // The index for the new row (0-based)

        // Create label and input
        const spellPowerOptions = this._mapStateToInputs().spellPowerProfiles.map(p => `<option value="${p.id}">SP ${p.id}</option>`).join('');
        const newAdditionalScalingId = existingScalingRows + 1; // A unique ID for the new elements
        const wrapper = document.createElement('div');
        wrapper.className = 'input-group-row additional-scaling-row';
        wrapper.innerHTML = `
            <label for="additional-scaling-base-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" class="short-label">SP ${newScalingIndex + 2} Base</label>
            <input type="text" data-element="additional-scaling-base-${spellRowId}-${newAdditionalScalingId}" id="additional-scaling-base-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Base damage for this component (does not scale with CL)">
            <span class="plus-symbol">+</span>
            <label for="additional-scaling-cl-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" class="short-label">per CL</label>
            <input type="text" data-element="additional-scaling-cl-${spellRowId}-${newAdditionalScalingId}" id="additional-scaling-cl-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Bonus damage per caster level for this component">
            <select data-element="additional-scaling-sp-select-${spellRowId}-${newAdditionalScalingId}" id="additional-scaling-sp-select-${spellRowId}-${newAdditionalScalingId}${this.idSuffix}" class="small-input" title="Select Spell Power source for this component">${spellPowerOptions}</select>            
            <button class="remove-scaling-input-btn small-btn" title="Remove this scaling input">&times;</button>
        `;

        scalingsContainer.appendChild(wrapper);

        // Set the default selection for the new dropdown
        const spSelect = wrapper.querySelector('select');
        if (spSelect) {
            // Default to the corresponding SP source if it exists, otherwise the last one.
            const targetProfileId = newScalingIndex + 2;
            const profiles = this._mapStateToInputs().spellPowerProfiles;
            if (profiles.some(p => p.id === targetProfileId)) {
                spSelect.value = targetProfileId;
            } else if (profiles.length > 0) {
                spSelect.value = profiles[profiles.length - 1].id;
            }
        }

        // If 5 inputs are present, hide the add button
        if (existingScalingRows + 1 >= 5) {
            mainRow.querySelector('.add-scaling-input-btn').classList.add('hidden');
        }

        // Add event listener for the new remove button
        wrapper.querySelector('.remove-scaling-input-btn').addEventListener('click', (e) => {
            e.preventDefault();
            wrapper.remove();
            // After removing, ensure the main row's add button is visible again if it was hidden
            const addBtn = mainRow.querySelector('.add-scaling-input-btn');
            if (addBtn) {
                addBtn.classList.remove('hidden');
            }

            this.handleInputChange(e);
        });

        // Recalculate damage and save state
        this.calculateSpellDamage();
        this.manager.saveState();
    }
    
    removeSpellDamageRow(sourceGroup) {
        const mainRow = sourceGroup.querySelector('.spell-damage-source-row');
        const rowId = mainRow.dataset.rowId;
        const rowData = {};
        sourceGroup.querySelectorAll('input[data-element], select[data-element]').forEach(input => {
            rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
        });

        const parent = sourceGroup.parentNode;
        const rowIndex = Array.prototype.indexOf.call(parent.children, sourceGroup);

        this.manager.recordAction({ type: 'REMOVE_DYNAMIC_ROW', setId: this.setId, rowType: 'spellDamage', rowId, rowData, rowIndex });

        sourceGroup.remove();

        // Simulate a change to trigger recalculation and state saving
        const fakeEvent = { target: this.container };
        this.handleInputChange(fakeEvent);
    }


    addSpellDamageRow(isProgrammatic = false) {
        let maxRowId = 0;
        this.spellDamageRowsContainer.querySelectorAll('.spell-damage-source-row').forEach(row => { // This still works as it finds all main rows
            const firstInput = row.querySelector('input[data-element^="spell-name-"]');
            if (firstInput) {
                const idNum = parseInt(firstInput.dataset.element.match(/(\d+)$/)[1], 10);
                if (idNum > maxRowId) {
                    maxRowId = idNum;
                }
            }
        });
        const newRowId = maxRowId + 1;

        const newGroup = document.createElement('div');
        newGroup.className = 'spell-source-group';
        newGroup.innerHTML = `
            <div class="input-group-row spell-damage-source-row" data-row-id="${newRowId}">
                <input type="text" data-element="spell-name-${newRowId}" id="spell-name-${newRowId}${this.idSuffix}" value="Source ${newRowId}" title="Name of the spell component" placeholder="Spell Name" class="adaptive-text-input">
                <label for="spell-damage-${newRowId}${this.idSuffix}">Base Damage</label>
                <input type="text" data-element="spell-damage-${newRowId}" id="spell-damage-${newRowId}${this.idSuffix}" value="0" title="The spell's base damage (e.g., 10d6+50)" class="adaptive-text-input">
                <span class="plus-symbol">+</span>
                <label for="spell-cl-scaling-${newRowId}${this.idSuffix}" class="short-label">per CL</label>
                <input type="text" data-element="spell-cl-scaling-${newRowId}" id="spell-cl-scaling-${newRowId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Bonus damage dice per caster level (e.g., 1d6 per CL)">
                <label for="caster-level-${newRowId}${this.idSuffix}" class="short-label">CL</label>
                <input type="number" data-element="caster-level-${newRowId}" id="caster-level-${newRowId}${this.idSuffix}" value="20" class="small-input" title="Caster Level for this damage component">
                <label for="spell-hit-count-${newRowId}${this.idSuffix}" class="short-label">Hits</label>
                <input type="number" data-element="spell-hit-count-${newRowId}" id="spell-hit-count-${newRowId}${this.idSuffix}" value="1" min="1" class="small-input" title="Number of times this spell component hits">
                <button class="add-scaling-input-btn small-btn" title="Add additional scaling input">+</button>
                <button class="duplicate-row-btn small-btn" title="Duplicate this damage source">❐</button>
                <button class="remove-row-btn" title="Remove this damage source">&times;</button>
            </div>
            <div class="additional-scalings-container"></div>
        `;
        this.spellDamageRowsContainer.appendChild(newGroup);
        newGroup.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));

        const rowData = {};
        newGroup.querySelectorAll('input[data-element]').forEach(input => {
            rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
        });
        const parent = newGroup.parentNode;
        const rowIndex = Array.prototype.indexOf.call(parent.children, newGroup);

        if (!isProgrammatic) {
            this.manager.recordAction({ type: 'ADD_DYNAMIC_ROW', setId: this.setId, rowType: 'spellDamage', rowId: newRowId, rowData, rowIndex });
            this.handleInputChange({ target: newGroup.querySelector('input') });
        }

        return newGroup; // Return the newly created group element
    }

    duplicateSpellDamageRow(originalGroup) {
        if (!originalGroup) return;

        const originalRow = originalGroup.querySelector('.spell-damage-source-row');
        if (!originalRow) return;

        // 1. Gather data from the original main row
        const originalData = {
            name: originalRow.querySelector('input[data-element^="spell-name-"]').value,
            base: originalRow.querySelector('input[data-element^="spell-damage-"]').value,
            clScaled: originalRow.querySelector('input[data-element^="spell-cl-scaling-"]').value,
            casterLevel: originalRow.querySelector('input[data-element^="caster-level-"]').value,
            hitCount: originalRow.querySelector('input[data-element^="spell-hit-count-"]').value,
        };

        // 2. Gather data from associated additional scaling rows
        const additionalScalingsData = [];
        originalGroup.querySelectorAll('.additional-scaling-row').forEach(scalingRow => {
            additionalScalingsData.push({
                base: scalingRow.querySelector('input[data-element^="additional-scaling-base-"]').value,
                clScaled: scalingRow.querySelector('input[data-element^="additional-scaling-cl-"]').value,
                profileId: scalingRow.querySelector('select[data-element^="additional-scaling-sp-select-"]').value,
            });
        });

        // 3. Create a new row
        const newRow = this.addSpellDamageRow(true); // Call programmatically

        // 4. Populate the new main row with the original data
        newRow.querySelector('input[data-element^="spell-name-"]').value = originalData.name + " (Copy)";
        newRow.querySelector('input[data-element^="spell-damage-"]').value = originalData.base;
        newRow.querySelector('input[data-element^="spell-cl-scaling-"]').value = originalData.clScaled;
        newRow.querySelector('input[data-element^="caster-level-"]').value = originalData.casterLevel;
        newRow.querySelector('input[data-element^="spell-hit-count-"]').value = originalData.hitCount;

        // 5. Re-create and populate the additional scaling rows for the new main row
        additionalScalingsData.forEach(scalingData => {
            this._addAdditionalScalingInput(newRow); // Create a new scaling row attached to the new group
            const lastScalingRow = Array.from(newRow.querySelectorAll('.additional-scaling-row')).pop();
            if (lastScalingRow) {
                lastScalingRow.querySelector('input[data-element^="additional-scaling-base-"]').value = scalingData.base;
                lastScalingRow.querySelector('input[data-element^="additional-scaling-cl-"]').value = scalingData.clScaled;
                lastScalingRow.querySelector('select[data-element^="additional-scaling-sp-select-"]').value = scalingData.profileId;
            }
        });

        // 6. Resize inputs and trigger a full recalculation
        this.resizeAllAdaptiveInputs();
        this.handleInputChange({ target: newRow.querySelector('.spell-damage-source-row input') });
        this.manager.saveState();
    }

    _mapStateToInputs() {
        const spellPowerProfiles = [];
        const spellDamageSources = [];

        // Iterate over state to build profiles and sources
        for (const key in this.state) {
            if (key.startsWith('spell-power-type-')) {
                const id = key.substring('spell-power-type-'.length);
                spellPowerProfiles.push({
                    id: parseInt(id, 10),
                    type: this.state[key],
                    spellPower: parseInt(this.state[`spell-power-${id}`], 10) || 0,
                    critChance: (parseFloat(this.state[`spell-crit-chance-${id}`]) || 0) / 100,
                    critDamage: (parseFloat(this.state[`spell-crit-damage-${id}`]) || 0) / 100,
                });
            } else if (key.startsWith('spell-name-')) {
                const id = key.substring('spell-name-'.length);
                const source = {
                    id: parseInt(id, 10),
                    name: this.state[key] || `Source ${id}`,
                    base: parseDiceNotation(this.state[`spell-damage-${id}`]),
                    clScaled: parseDiceNotation(this.state[`spell-cl-scaling-${id}`]),
                    casterLevel: parseInt(this.state[`caster-level-${id}`], 10) || 0,
                    hitCount: parseInt(this.state[`spell-hit-count-${id}`], 10) || 1,
                    additionalScalings: []
                };

                // Find additional scalings for this source
                for (const subKey in this.state) {
                    if (subKey.startsWith(`additional-scaling-base-${id}-`)) {
                        const scalingId = subKey.substring(`additional-scaling-base-${id}-`.length);
                        source.additionalScalings.push({
                            base: parseDiceNotation(this.state[subKey]),
                            clScaled: parseDiceNotation(this.state[`additional-scaling-cl-${id}-${scalingId}`]),
                            profileId: parseInt(this.state[`additional-scaling-sp-select-${id}-${scalingId}`], 10) || 1,
                        });
                    }
                }
                spellDamageSources.push(source);
            }
        }

        return {
            spellDamageSources,
            spellPowerProfiles,
            isEmpowered: this.state['metamagic-empower'],
            isMaximized: this.state['metamagic-maximize'],
            isIntensified: this.state['metamagic-intensify'],
            isWellspring: this.state['boost-wellspring'],
            isNightHorrors: this.state['boost-night-horrors'],
        };
    }

    handleInputChange(e) {
        const input = e.target;
        const key = input.dataset.element;
        if (!key) return;
        
        let value;
        if (input.type === 'checkbox') {
            value = input.checked;
        } else if (input.type === 'number') {
            value = input.value === '' ? 0 : parseFloat(input.value);
        } else {
            value = input.value;
        }

        this.state[key] = value;

        this.calculateSpellDamage();
        this.manager.updateComparisonTable();
        this.manager.saveState();
    }

    calculateSpellDamage() {
        let totalBaseDamage = 0;
        const individualSpellDamages = []; // To store results for each spell
        const inputs = this._mapStateToInputs();

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

        let totalBoostCritDamageBonus = 0;
        const wellspringSpellPowerBonus = inputs.isWellspring ? 150 : 0;
        if (inputs.isWellspring) {
            totalBoostCritDamageBonus += 0.20;
        }
        if (inputs.isNightHorrors) {
            totalBoostCritDamageBonus += 0.25;
        }
        const totalSpellPowerBonus = metamagicSpellPower + wellspringSpellPowerBonus;
 
        // Update bonus displays for each spell power profile
        this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile').forEach(profileEl => {
            const profileId = profileEl.dataset.profileId;
            const spellPowerBonusSpan = profileEl.querySelector(`#metamagic-spell-power-bonus-${profileId}${this.idSuffix}`);
            const critDamageBonusSpan = profileEl.querySelector(`#boost-crit-damage-bonus-${profileId}${this.idSuffix}`);
            if (spellPowerBonusSpan) spellPowerBonusSpan.textContent = totalSpellPowerBonus;
            if (critDamageBonusSpan) critDamageBonusSpan.textContent = `${totalBoostCritDamageBonus * 100}%`;
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
                const critMultiplier = 2 + baseProfile.critDamage + totalBoostCritDamageBonus;

                const averageHit = baseDamageComponent * spellPowerMultiplier;
                const averageCrit = averageHit * critMultiplier;
                const totalAverage = (averageHit * (1 - baseProfile.critChance)) + (averageCrit * baseProfile.critChance);

                sourceTotalAverage += totalAverage; // Keep track of the running total for the whole source
                components.push({
                    name: `SP 1: ${baseProfile.type || 'Unnamed'}`,
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
                    const critMultiplier = 2 + profile.critDamage + totalBoostCritDamageBonus;

                    // Combine base and CL-scaled damage for this component
                    const damageComponent = scaling.base + (scaling.clScaled * source.casterLevel);
                    const averageHit = damageComponent * spellPowerMultiplier;
                    const averageCrit = averageHit * critMultiplier;
                    const totalAverage = (averageHit * (1 - profile.critChance)) + (averageCrit * profile.critChance);

                    sourceTotalAverage += totalAverage; // Add to the running total
                    components.push({
                        name: `SP ${scaling.profileId}: ${profile.type || 'Unnamed'}`,
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

        let totalAverageBaseHit = 0;
        let totalAverageCritHit = 0;
        individualSpellDamages.forEach(spell => {
            spell.components.forEach(component => {
                totalAverageBaseHit += component.averageHit;
                totalAverageCritHit += component.averageCrit;
            });
        });

        this.averageBaseHit = totalAverageBaseHit;
        this.averageCritHit = totalAverageCritHit;

        // Update UI
        this.avgSpellDamageSpan.textContent = totalBaseDamage.toFixed(2); // This now represents total pre-MRR average
        this.totalAvgSpellDamageSpan.textContent = totalBaseDamage.toFixed(2);
        this.finalSpellDamageSpan.textContent = totalBaseDamage.toFixed(2);

        this._updateSummaryUI(); // Call the new UI update method
        this.updateSpellPowerSelectors();
    }

    _updateSummaryUI() {
        if (!this.individualSpellDamageSummary) return;

        this.individualSpellDamageSummary.innerHTML = '<h3>Individual Spell Damage</h3>'; // Clear previous results

        (this.individualSpellDamages || []).forEach(spell => {
            const spellContainer = document.createElement('div');
            spellContainer.style.marginBottom = '0.5rem';
            spellContainer.innerHTML = `<strong>${spell.name} (Total Avg: ${spell.totalAverage.toFixed(2)})</strong>`;

            spell.components.forEach(component => {
                const p = document.createElement('p');
                p.style.marginLeft = '1rem';
                p.innerHTML = `<em>${component.name}</em><br>Avg Hit: ${component.averageHit.toFixed(2)}, Avg Crit: ${component.averageCrit.toFixed(2)}`;
                spellContainer.appendChild(p);
            });
            this.individualSpellDamageSummary.appendChild(spellContainer);
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
        const chanceValue = state?.critChance * 100 || '20';
        const damageValue = state?.critDamage * 100 || '100';

        newProfile.innerHTML = `
            <div class="profile-header">
                <label class="section-headline">Spell Power & Criticals ${newProfileId}</label>
                <input type="text" data-element="spell-power-type-${newProfileId}" id="spell-power-type-${newProfileId}${this.idSuffix}" class="adaptive-text-input spell-power-type-input"
                    placeholder="Element Type" title="e.g., Fire, Acid, etc." value="${typeValue}">
            </div>
            <div class="input-group-row">
                <label for="spell-power-${newProfileId}${this.idSuffix}">Spell Power</label>
                <input type="number" data-element="spell-power-${newProfileId}" id="spell-power-${newProfileId}${this.idSuffix}" value="${powerValue}" class="small-input" title="Your spell power for this profile.">
                <span class="plus-symbol">+</span>
                <span id="metamagic-spell-power-bonus-${newProfileId}${this.idSuffix}" class="read-only-bonus">0</span>
            </div>
            <div class="input-group-row">
                <label for="spell-crit-chance-${newProfileId}${this.idSuffix}">Crit Chance %</label>
                <input type="number" data-element="spell-crit-chance-${newProfileId}" id="spell-crit-chance-${newProfileId}${this.idSuffix}" value="${chanceValue}" class="small-input" title="Your chance to critically hit with this profile.">
            </div>
            <div class="input-group-row">
                <label for="spell-crit-damage-${newProfileId}${this.idSuffix}">Crit Dmg Bonus %</label>
                <input type="number" data-element="spell-crit-damage-${newProfileId}" id="spell-crit-damage-${newProfileId}${this.idSuffix}" value="${damageValue}" class="small-input" title="Your additional spell critical damage bonus. Base critical damage is +100% (x2 total).">
                <span class="plus-symbol">+</span>
                <span id="boost-crit-damage-bonus-${newProfileId}${this.idSuffix}" class="read-only-bonus">0</span>
            </div>
            <button class="remove-row-btn" title="Remove this profile">&times;</button>
        `;

        this.spellPowerSourcesContainer.appendChild(newProfile);

        // Only trigger recalculation and save if we are NOT in a bulk-loading state (i.e., state is null)
        if (state === null) {
            // Add new profile to state
            this.state[`spell-power-type-${newProfileId}`] = typeValue;
            this.state[`spell-power-${newProfileId}`] = powerValue;
            this.state[`spell-crit-chance-${newProfileId}`] = chanceValue;
            this.state[`spell-crit-damage-${newProfileId}`] = damageValue;

            const rowData = {};
            newProfile.querySelectorAll('input[data-element]').forEach(input => {
                rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
            });
            const parent = newProfile.parentNode;
            const rowIndex = Array.prototype.indexOf.call(parent.children, newProfile);

            this.manager.recordAction({ type: 'ADD_DYNAMIC_ROW', setId: this.setId, rowType: 'spellPower', rowId: newProfileId, rowData, rowIndex });

            this.updateSpellPowerSelectors();
            this.handleInputChange({ target: newProfile.querySelector('input') }); // Simulate an input change
        }
    }

    removeSpellPowerSource(profileElement) {
        if (this.spellPowerSourcesContainer.children.length <= 1) {
            alert("You cannot remove the last spell power source.");
            return;
        }
        const profileId = profileElement.dataset.profileId;

        const rowData = {};
        profileElement.querySelectorAll('input[data-element]').forEach(input => {
            rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
        });

        const parent = profileElement.parentNode;
        const rowIndex = Array.prototype.indexOf.call(parent.children, profileElement);

        this.manager.recordAction({ type: 'REMOVE_DYNAMIC_ROW', setId: this.setId, rowType: 'spellPower', rowId: profileId, rowData, rowIndex });

        // Remove from state
        delete this.state[`spell-power-type-${profileId}`];
        delete this.state[`spell-power-${profileId}`];
        delete this.state[`spell-crit-chance-${profileId}`];
        delete this.state[`spell-crit-damage-${profileId}`];

        profileElement.remove();
        this.updateSpellPowerSelectors();
        this.handleInputChange({ target: this.container }); // Simulate a change to trigger recalc and save
    }

    updateSpellPowerSelectors() {
        const profiles = this._mapStateToInputs().spellPowerProfiles;
        const profileOptions = profiles.map(p => `<option value="${p.id}">SP ${p.id}</option>`).join('');

        this.container.querySelectorAll(`select[data-element^="additional-scaling-sp-select-"]`).forEach(select => {
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

    setState(state) {
        // Clear dynamic rows before setting state
        this.spellDamageRowsContainer.querySelectorAll('.spell-damage-source-row:not(:first-child)').forEach(row => row.remove());
        this.spellPowerSourcesContainer.querySelectorAll('.spell-power-profile:not(:first-child)').forEach(row => row.remove());

        this.state = { ...this.state, ...state };

        // Re-create dynamic rows from the new state
        for (const key in this.state) {
            if (key.startsWith('spell-name-') && key !== 'spell-name-1') {
                // This is a bit tricky, we need to create the row but not add it to state again
                this.addSpellDamageRow(true);
            } else if (key.startsWith('spell-power-type-') && key !== 'spell-power-type-1') {
                const id = key.substring('spell-power-type-'.length);
                const profileState = {
                    id: id,
                    type: this.state[key],
                    spellPower: this.state[`spell-power-${id}`],
                    critChance: this.state[`spell-crit-chance-${id}`] / 100,
                    critDamage: this.state[`spell-crit-damage-${id}`] / 100,
                };
                this.addSpellPowerSource(profileState);
            }
        }

        super.setState(this.state); // This calls deserializeForm
        this.calculateSpellDamage();
    }

    removeDynamicRow(rowId, rowType) {
        let container, selector;
        if (rowType === 'spellDamage') {
            container = this.spellDamageRowsContainer;
            selector = `.spell-damage-source-row[data-row-id="${rowId}"]`;
        } else if (rowType === 'spellPower') {
            container = this.spellPowerSourcesContainer;
            selector = `.spell-power-profile[data-profile-id="${rowId}"]`;
        } else {
            return;
        }

        const row = container.querySelector(selector);
        if (row) {
            row.remove();
            this.calculateSpellDamage();
            this.manager.saveState();
        }
    }

    recreateDynamicRow(rowId, rowType, rowData, rowIndex) {
        let newRow;
        let container;

        if (rowType === 'spellDamage') {
            container = this.spellDamageRowsContainer;
            newRow = document.createElement('div'); // Create the outer group
            newRow.className = 'spell-source-group';
            newRow.innerHTML = `
                <div class="input-group-row spell-damage-source-row" data-row-id="${rowId}">
                    <input type="text" data-element="spell-name-${rowId}" id="spell-name-${rowId}${this.idSuffix}" value="Source ${rowId}" title="Name of the spell component" placeholder="Spell Name" class="adaptive-text-input">
                    <label for="spell-damage-${rowId}${this.idSuffix}">Base Damage</label>
                    <input type="text" data-element="spell-damage-${rowId}" id="spell-damage-${rowId}${this.idSuffix}" value="0" title="The spell's base damage (e.g., 10d6+50)" class="adaptive-text-input">
                    <span class="plus-symbol">+</span>
                    <label for="spell-cl-scaling-${rowId}${this.idSuffix}" class="short-label">per CL</label>
                    <input type="text" data-element="spell-cl-scaling-${rowId}" id="spell-cl-scaling-${rowId}${this.idSuffix}" value="0" class="small-input adaptive-text-input" title="Bonus damage dice per caster level (e.g., 1d6 per CL)">
                    <label for="caster-level-${rowId}${this.idSuffix}" class="short-label">CL</label>
                    <input type="number" data-element="caster-level-${rowId}" id="caster-level-${rowId}${this.idSuffix}" value="20" class="small-input" title="Caster Level for this damage component">
                    <label for="spell-hit-count-${rowId}${this.idSuffix}" class="short-label">Hits</label>
                    <input type="number" data-element="spell-hit-count-${rowId}" id="spell-hit-count-${rowId}${this.idSuffix}" value="1" min="1" class="small-input" title="Number of times this spell component hits">
                    <button class="add-scaling-input-btn small-btn" title="Add additional scaling input">+</button>
                    <button class="duplicate-row-btn small-btn" title="Duplicate this damage source">❐</button>
                    <button class="remove-row-btn" title="Remove this damage source">&times;</button>
                </div>
                <div class="additional-scalings-container"></div>
            `;
        } else if (rowType === 'spellPower') {
            this.addSpellPowerSource(rowData); // This is simpler as it can recreate from state-like object
            return; // addSpellPowerSource handles insertion and calculation
        } else {
            return;
        }

        // Populate the state for the newly created row
        Object.keys(rowData).forEach(key => {
            this.state[key] = rowData[key];
        });

        container.insertBefore(newRow, container.children[rowIndex]);
        this.setState(this.state); // Full refresh
    }
}
