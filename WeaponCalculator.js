import { parseDiceNotation } from './utils.js';

export class WeaponCalculator {
        constructor(setId, manager, name) { // Add 'name' to the constructor
            this.setId = setId;
            // Properties to store calculation results for the comparison table
            this.totalAverageDamage = 0;
            this.totalAvgBaseHitDmg = 0;
            this.totalAvgSneakDmg = 0;
            this.totalAvgUnscaledDmg = 0; // New property for unscaled damage
            this.totalAvgImbueDmg = 0; // This will hold the sum of all unscaled sources
            this.totalAvgScaledDiceDmg = 0;

            this.idSuffix = `-set${setId}`;
            this.getElements();
            this.manager = manager;
            this.addEventListeners();
            // Dynamically count the number of unscaled rows present in the template/DOM.
            this.unscaledRowCount = this.unscaledRowsContainer.querySelectorAll('.input-group-row').length;

            this.recalculateHandler = this.handleInputChange.bind(this);
            // We will call the initial calculation from the manager

            // Create a hidden span for measuring text width
            this._measurementSpan = document.createElement('span');
            this._measurementSpan.style.position = 'absolute';
            this._measurementSpan.style.visibility = 'hidden';
            this._measurementSpan.style.whiteSpace = 'nowrap';
            document.body.appendChild(this._measurementSpan);

            // Initialize adaptive sizing for all relevant inputs
            this._initializeAdaptiveInputs();
        }

        /**
         * Resizes an individual input element to fit its content.
         * @param {HTMLInputElement} inputElement - The input element to resize.
         */
        _resizeInput(inputElement) {
            // Apply relevant styles from the input to the measurement span
            const computedStyle = window.getComputedStyle(inputElement);
            this._measurementSpan.style.fontFamily = computedStyle.fontFamily;
            this._measurementSpan.style.fontSize = computedStyle.fontSize;
            this._measurementSpan.style.fontWeight = computedStyle.fontWeight;
            this._measurementSpan.style.letterSpacing = computedStyle.letterSpacing;
            this._measurementSpan.style.textTransform = computedStyle.textTransform;
            // Add padding from the input, but be careful with box-sizing
            const paddingLeft = parseFloat(computedStyle.paddingLeft);
            const paddingRight = parseFloat(computedStyle.paddingRight);
            const borderWidthLeft = parseFloat(computedStyle.borderLeftWidth);
            const borderWidthRight = parseFloat(computedStyle.borderRightWidth);

            this._measurementSpan.textContent = inputElement.value || inputElement.placeholder || '';

            // Calculate the desired width including padding and border
            // Adding a small buffer (e.g., 2-4px) to prevent scrollbars from appearing prematurely.
            let desiredWidth = this._measurementSpan.offsetWidth + paddingLeft + paddingRight + borderWidthLeft + borderWidthRight + 4;

            const minWidth = parseFloat(computedStyle.minWidth) || 50; // Use min-width from CSS or default
            // If the input is in a flex container (which .input-group-row is),
            // its max-width might be constrained by the container.
            // For now, let's cap it at a reasonable value or its parent's width.
            let maxWidth = parseFloat(computedStyle.maxWidth);
            if (isNaN(maxWidth) || maxWidth === 0) { // If CSS max-width is 'none' or '0px', use a default.
                maxWidth = 150; // A reasonable default maximum width
            }

            inputElement.style.width = `${Math.min(maxWidth, Math.max(minWidth, desiredWidth))}px`;
        }

        /**
         * Initializes adaptive sizing for all relevant inputs in the calculator.
         * Uses event delegation for dynamically added inputs.
         */
        _initializeAdaptiveInputs() {
            const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
            if (!calculatorElement) return;

            // Use event delegation for 'input' events to handle resizing
            calculatorElement.addEventListener('input', (e) => {
                // Only resize inputs that have the 'adaptive-text-input' class
                if (e.target.classList.contains('adaptive-text-input')) {
                    this._resizeInput(e.target);
                }
            });

            // Run initial resize on all existing adaptive inputs
            this.resizeAllAdaptiveInputs();
        }

        /**
         * Resizes all input elements that have the 'adaptive-text-input' class.
         */
        resizeAllAdaptiveInputs() {
            const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
            if (!calculatorElement) return;
            calculatorElement.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
        }

        getElements() {
            const get = (baseId) => document.getElementById(baseId + this.idSuffix);

            this.weaponDiceInput = get('weapon-dice'); // For setId=1, suffix is '', gets 'weapon-dice'. For setId=2, suffix is '-set2', gets 'weapon-dice-set2'.
            this.weaponDamageInput = get('weapon-damage');
            this.bonusBaseDamageInput = get('bonus-base-damage');
            this.meleePowerInput = get('melee-power');
            this.spellPowerInput = get('spell-power');
            this.critThreatInput = get('crit-threat');
            this.critMultiplierInput = get('crit-multiplier');
            this.seekerDamageInput = get('seeker-damage');
            this.critMultiplier1920Input = get('crit-multiplier-19-20');
            this.sneakAttackDiceInput = get('sneak-attack-dice');
            this.sneakBonusInput = get('sneak-bonus');
            this.doublestrikeInput = get('doublestrike');
            this.isDoubleshotCheckbox = get('is-doubleshot');
            this.missThresholdInput = get('miss-threshold');
            this.grazeThresholdInput = get('graze-threshold');
            this.reaperSkullsSelect = get('reaper-skulls');
            this.grazePercentInput = get('graze-percent');
            this.imbueDiceCountInput = get('imbue-dice-count');
            this.imbueDieTypeInput = get('imbue-die-type');
            this.imbueScalingInput = get('imbue-scaling');
            this.imbueUsesSpellpowerCheckbox = get('imbue-uses-spellpower');
            this.imbueCritsCheckbox = get('imbue-crits');
            this.imbueToggleBonusSpan = get('imbue-toggle-bonus');
            this.calculateBtn = get('calculate-btn');
            this.avgBaseDamageSpan = get('avg-base-damage');
            this.avgSneakDamageSpan = get('avg-sneak-damage');
            this.avgImbueDamageSpan = get('avg-imbue-damage');
            this.avgUnscaledDamageSpan = get('avg-unscaled-damage'); // Get the new summary span
            this.avgScaledDiceDamageSpan = get('avg-scaled-dice-damage');
            this.totalAvgDamageSpan = get('total-avg-damage');
            this.weaponScalingSpan = get('weapon-scaling');
            this.sneakScalingSpan = get('sneak-scaling');
            this.imbueScalingBreakdownSpan = get('imbue-scaling-breakdown');
            this.scaledDiceScalingBreakdownSpan = get('scaled-dice-scaling-breakdown');
            this.scaledDiceAddedInputDisplaySpan = get('scaled-dice-added-input-display'); // New display in input area
            this.imbuePowerSourceSpan = get('imbue-power-source');
            this.summaryHeader = get('summary-header');
            this.reaperPenaltySpan = get('reaper-penalty');
            this.rollDamageTbody = get('roll-damage-tbody');

            // Preset buttons
            this.set75ScalingBtn = get('set-75-scaling-btn');
            this.set100ScalingBtn = get('set-100-scaling-btn');
            this.set150ScalingBtn = get('set-150-scaling-btn');
            this.set200ScalingBtn = get('set-200-scaling-btn');

            this.unscaledRowsContainer = get('unscaled-rows-container');
            this.addUnscaledRowBtn = get('add-unscaled-row-btn');

            // Scaled Dice Damage elements
            this.scaledDiceRowsContainer = get('scaled-dice-rows-container');
            this.addScaledDiceRowBtn = get('add-scaled-dice-row-btn');
        }

        /**
         * Parses the critical threat input, which can be a number (e.g., "5")
         * or a range (e.g., "16-20").
         * @param {string} threatString - The value from the crit threat input field.
         * @returns {number} The size of the threat range.
         */
        parseThreatRange(threatString) {
            const cleanString = (threatString || '').trim();

            if (cleanString.includes('-')) {
                const parts = cleanString.split('-');
                if (parts.length === 2) {
                    const start = parseInt(parts[0], 10);
                    const end = parseInt(parts[1], 10);
                    // Check for valid numbers and a valid DDO range (e.g., 16-20, not 20-16)
                    if (!isNaN(start) && !isNaN(end) && start <= end && start >= 1 && end <= 20) {
                        return end - start + 1;
                    }
                }
            } else {
                const rangeSize = parseInt(cleanString, 10);
                if (!isNaN(rangeSize) && rangeSize >= 1 && rangeSize <= 20) {
                    return rangeSize;
                }
            }
            return 1; // Default to a range of 1 (a roll of 20) if input is invalid
        }

        /**
         * Gathers all raw input values from the DOM.
         * @returns {object} An object containing all necessary input values for calculation.
         */
        _getInputs() {
            const isDoubleshot = this.isDoubleshotCheckbox.checked;
            let multiStrikeValue = parseFloat(this.doublestrikeInput.value) || 0;
            if (!isDoubleshot) {
                multiStrikeValue = Math.min(multiStrikeValue, 100);
            }

            const unscaled = {
                normal_multi: 0,
                normal_noMulti: 0,
                crit_multi: 0,
                crit_noMulti: 0
            };

            const rows = this.unscaledRowsContainer.querySelectorAll('.input-group-row');

            rows.forEach(row => {
                const dmgInput = row.querySelector(`input[id^="unscaled-damage-"]`);
                const procInput = row.querySelector(`input[id^="unscaled-proc-chance-"]`);
                const multiStrikeCheckbox = row.querySelector(`input[id^="unscaled-doublestrike-"]`);
                const onCritCheckbox = row.querySelector(`input[id^="unscaled-on-crit-"]`);

                if (dmgInput && procInput && multiStrikeCheckbox && onCritCheckbox) {
                    const damage = parseDiceNotation(dmgInput.value);
                    const procChance = (parseFloat(procInput.value) || 100) / 100;
                    const averageDamage = damage * procChance;

                    if (multiStrikeCheckbox.checked) {
                        if (onCritCheckbox.checked) unscaled.crit_multi += averageDamage;
                        else unscaled.normal_multi += averageDamage;
                    } else {
                        if (onCritCheckbox.checked) unscaled.crit_noMulti += averageDamage;
                        else unscaled.normal_noMulti += averageDamage;
                    }
                }
            });


            const scaledDiceDamage = [];
            this.scaledDiceRowsContainer.querySelectorAll('.input-group-row').forEach(row => {
                const baseDiceInput = row.querySelector(`input[id^="scaled-dice-base-"]`);
                const scalingToggle = row.querySelector(`input[id^="scaled-dice-scaling-toggle-"]`);
                const enabledCheckbox = row.querySelector(`input[id^="scaled-dice-enabled-"]`);

                // Only process the row if it's enabled.
                if (baseDiceInput && scalingToggle && enabledCheckbox && enabledCheckbox.checked) {
                    const procChanceInput = row.querySelector(`input[id^="scaled-dice-proc-chance-"]`);
                    scaledDiceDamage.push({
                        baseDice: baseDiceInput.value,
                        procChance: (parseFloat(procChanceInput.value) || 100) / 100,
                        enableScaling: scalingToggle.checked,
                        scalingPercent: parseFloat(row.querySelector(`input[id^="scaled-dice-scaling-percent-"]`).value) || 100,
                        isEnabled: enabledCheckbox.checked
                    });
                }
            });

            return {
                additionalWeaponDice: parseFloat(this.weaponDiceInput.value) || 0,
                parsedWeaponDmg: parseDiceNotation(this.weaponDamageInput.value) || 0,
                bonusBaseDmg: parseFloat(this.bonusBaseDamageInput.value) || 0,
                meleePower: parseFloat(this.meleePowerInput.value) || 0,
                spellPower: parseFloat(this.spellPowerInput.value) || 0,
                threatRange: this.parseThreatRange(this.critThreatInput.value),
                critMult: parseFloat(this.critMultiplierInput.value) || 2,
                critMult1920: parseFloat(this.critMultiplier1920Input.value) || 0,
                seekerDmg: parseDiceNotation(this.seekerDamageInput.value), // Allow dice notation for seeker
                sneakDiceCount: parseInt(this.sneakAttackDiceInput.value) || 0,
                sneakBonusDmg: parseFloat(this.sneakBonusInput.value) || 0,
                missThreshold: Math.max(1, parseInt(this.missThresholdInput.value) || 1),
                grazeThreshold: parseInt(this.grazeThresholdInput.value) || 0,
                grazePercent: (parseFloat(this.grazePercentInput.value) || 0) / 100,
                reaperSkulls: parseInt(this.reaperSkullsSelect.value) || 0,
                imbueDiceCount: parseInt(this.imbueDiceCountInput.value) || 0,
                imbueDieType: parseInt(this.imbueDieTypeInput.value) || 6,
                imbueScaling: (parseFloat(this.imbueScalingInput.value) || 100) / 100,
                imbueCrits: this.imbueCritsCheckbox.checked,
                imbueUsesSpellpower: this.imbueUsesSpellpowerCheckbox.checked,
                doublestrikeChance: multiStrikeValue / 100,
                isDoubleshot: isDoubleshot,
                unscaled: unscaled,
                scaledDiceDamage: scaledDiceDamage
            };
        }

        /**
         * Calculates the probabilities of different hit outcomes based on d20 rolls.
         * @param {object} inputs - The object containing miss, graze, and crit thresholds.
         * @returns {object} An object with probabilities for each outcome.
         */
        _calculateProbabilities(inputs) {
            const { missThreshold, grazeThreshold, threatRange } = inputs;
            const critStartRoll = 21 - threatRange;
            let missChance = 0, grazeChance = 0, normalCritChance = 0, specialCritChance = 0, normalChance = 0;

            for (let roll = 1; roll <= 20; roll++) {
                if (roll <= missThreshold) missChance++;
                else if (roll <= grazeThreshold) grazeChance++;
                else if (roll >= 19 && roll >= critStartRoll) specialCritChance++;
                else if (roll >= critStartRoll) normalCritChance++;
                else normalChance++;
            }

            return {
                miss: missChance / 20,
                graze: grazeChance / 20,
                normal: normalChance / 20,
                normalCrit: normalCritChance / 20,
                specialCrit: specialCritChance / 20,
                crit: (normalCritChance + specialCritChance) / 20,
                hit: (normalChance + normalCritChance + specialCritChance) / 20
            };
        }

        /**
         * Calculates the base damage values for each component before applying probabilities or multipliers.
         * @param {object} inputs - The object containing all raw input values.
         * @returns {object} An object containing the calculated damage portions.
         */
        _calculateDamagePortions(inputs) {
            const {
                additionalWeaponDice, parsedWeaponDmg, bonusBaseDmg, meleePower, spellPower,
                seekerDmg, sneakDiceCount, sneakBonusDmg, imbueDiceCount, imbueDieType, // This is a long line
                imbueScaling, imbueUsesSpellpower, scaledDiceDamage
            } = inputs;

            const baseDmg = parsedWeaponDmg + (additionalWeaponDice * parsedWeaponDmg) + bonusBaseDmg;
            const powerMultiplier = 1 + (meleePower / 100);
            const weaponPortion = baseDmg * powerMultiplier;
            const seekerPortion = seekerDmg * powerMultiplier;

            const sneakDiceDmg = sneakDiceCount * 3.5;
            const sneakPortion = (sneakDiceDmg + sneakBonusDmg) * (1 + (meleePower * 1.5) / 100);

            const imbueDice = imbueDiceCount * (imbueDieType + 1) / 2;
            const totalImbueDiceCount = imbueDiceCount > 0 ? imbueDiceCount + 1 : 0;
            const totalImbueDiceAverage = totalImbueDiceCount * (imbueDieType + 1) / 2;

            const powerForImbue = imbueUsesSpellpower ? spellPower : meleePower;
            const imbuePortion = totalImbueDiceAverage * (1 + (powerForImbue * imbueScaling) / 100);

            let totalAvgScaledDiceDmg = 0;
            let totalAddedScaledDice = 0;
            let scaledDiceBreakdownText = [];
            const addedDiceBreakdown = [];
            scaledDiceDamage.forEach(scaledDmg => {
                const baseDiceAvg = parseDiceNotation(scaledDmg.baseDice);
                const imbueThreshold = 7; // Fixed threshold
                const addDicePerThreshold = 1; // Fixed additional dice

                let additionalDiceAvg = 0;
                const scalingPercent = (scaledDmg.scalingPercent || 100) / 100;
                const powerMultiplier = 1 + (meleePower * scalingPercent / 100);

                if (scaledDmg.enableScaling) {
                    // How many times the threshold is met determines how many bonus dice are added.
                    const numAdditionalDice = Math.floor(imbueDiceCount / imbueThreshold) * addDicePerThreshold;
                    
                    // This is for the UI display. It should just be the number of dice added.
                    totalAddedScaledDice += numAdditionalDice;

                    if (numAdditionalDice > 0) {
                        addedDiceBreakdown.push(numAdditionalDice);
                    }
                    
                    // For damage calculation, we need the average of a single die of the specified type.
                    const baseDiceParts = scaledDmg.baseDice.toLowerCase().split('d');
                    if (baseDiceParts.length === 2) {
                        const dieType = baseDiceParts[1];
                        const singleDieAvg = parseDiceNotation(`1d${dieType}`);
                        additionalDiceAvg = numAdditionalDice * singleDieAvg;
                    }
                }
                // The total average from the dice rolls, which is then scaled by power.
                const totalDiceAverage = baseDiceAvg + additionalDiceAvg;
                const lineResult = (totalDiceAverage * powerMultiplier) * scaledDmg.procChance;
                totalAvgScaledDiceDmg += lineResult;

                // Build the text for this specific source
                let breakdown = `(${baseDiceAvg.toFixed(2)} (base) + ${additionalDiceAvg.toFixed(2)} (imbue)) * ${powerMultiplier.toFixed(2)} (power) * ${scaledDmg.procChance.toFixed(2)} (proc) = ${lineResult.toFixed(2)}`;
                scaledDiceBreakdownText.push(breakdown); 
            });

            const finalScaledDiceBreakdown = scaledDiceBreakdownText.join(' +<br>');

            return { baseDmg, weaponPortion, seekerPortion, sneakDiceDmg, sneakPortion, imbueDice: totalImbueDiceAverage, imbuePortion, powerForImbue, totalAvgScaledDiceDmg, addedDiceBreakdown, finalScaledDiceBreakdown, totalImbueDiceCount };
        }

        /**
         * Calculates the final average damage values by combining portions, probabilities, and multipliers.
         * @param {object} portions - The calculated damage portions.
         * @param {object} probabilities - The calculated outcome probabilities.
         * @param {object} inputs - The raw input values.
         * @returns {object} An object containing the final average damage for each component.
         */
        _calculateAverages(portions, probabilities, inputs) {
            const { critMult, critMult1920, grazePercent, reaperSkulls, doublestrikeChance, imbueCrits, unscaled } = inputs;
            const { weaponPortion, seekerPortion, sneakPortion, imbuePortion } = portions;

            const finalCritMult1920 = critMult + critMult1920;
            const multiStrikeMultiplier = 1 + doublestrikeChance;

            let reaperMultiplier = 1.0;
            if (reaperSkulls > 0) {
                if (reaperSkulls <= 6) reaperMultiplier = 20 / ((reaperSkulls ** 2) + reaperSkulls + 24);
                else reaperMultiplier = 5 / (4 * reaperSkulls - 8);
            }

            const avgBaseHitDmg =
                (((weaponPortion + seekerPortion) * finalCritMult1920) * probabilities.specialCrit) +
                (((weaponPortion + seekerPortion) * critMult) * probabilities.normalCrit) +
                (weaponPortion * probabilities.normal) +
                (weaponPortion * grazePercent * probabilities.graze);

            const avgSneakDmg = sneakPortion * (1 - probabilities.miss);
            const avgImbueDmg = imbuePortion * (probabilities.hit + (imbueCrits ? probabilities.crit : 0)); // Applies on hit, and extra on crit if checked

            // Correctly calculate unscaled damage based on probabilities
            const unscaledNormalPortion = (unscaled.normal_multi * multiStrikeMultiplier) + unscaled.normal_noMulti;
            const unscaledCritPortion = (unscaled.crit_multi * multiStrikeMultiplier) + unscaled.crit_noMulti;

            const avgUnscaledDmg =
                (unscaledNormalPortion * probabilities.hit) + // Normal unscaled damage on any non-graze, non-miss hit
                (unscaledCritPortion * probabilities.crit);   // "On Crit" unscaled damage only on crits

            // Correctly calculate scaled dice damage based on probabilities
            // It applies on any normal or critical hit, but not on grazes or misses.
            const avgScaledDiceDmg = portions.totalAvgScaledDiceDmg * probabilities.hit;

            return {
                base: avgBaseHitDmg * multiStrikeMultiplier * reaperMultiplier,
                sneak: avgSneakDmg * multiStrikeMultiplier * reaperMultiplier,
                imbue: avgImbueDmg * multiStrikeMultiplier * reaperMultiplier,
                unscaled: avgUnscaledDmg * reaperMultiplier, // Apply reaper penalty at the end
                scaledDice: avgScaledDiceDmg * multiStrikeMultiplier * reaperMultiplier,
                reaperMultiplier,
                multiStrikeMultiplier
            };
        }

        /**
         * Updates the summary and breakdown sections of the UI with calculated results.
         * @param {object} averages - The final average damage values.
         * @param {object} portions - The calculated damage portions.
         * @param {object} inputs - The raw input values.
         */
        _updateSummaryUI(averages, portions, inputs) {
            this.avgBaseDamageSpan.textContent = averages.base.toFixed(2);
            this.avgSneakDamageSpan.textContent = averages.sneak.toFixed(2);
            this.avgImbueDamageSpan.textContent = averages.imbue.toFixed(2);
            this.avgUnscaledDamageSpan.textContent = averages.unscaled.toFixed(2);
            this.avgScaledDiceDamageSpan.textContent = averages.scaledDice.toFixed(2);
            this.totalAvgDamageSpan.textContent = this.totalAverageDamage.toFixed(2);

            const { baseDmg, sneakDiceDmg, sneakPortion, imbueDice, imbuePortion, powerForImbue } = portions;
            const { meleePower, sneakBonusDmg, imbueScaling, imbueUsesSpellpower, spellPower } = inputs;

            this.weaponScalingSpan.textContent = `${baseDmg.toFixed(2)} * (1 + (${meleePower} / 100)) = ${portions.weaponPortion.toFixed(2)}`;
            this.sneakScalingSpan.textContent = `(${sneakDiceDmg.toFixed(2)} + ${sneakBonusDmg.toFixed(2)}) * (1 + (${meleePower} * 1.5) / 100) = ${sneakPortion.toFixed(2)}`;
            this.imbueScalingBreakdownSpan.textContent = `${portions.imbueDice.toFixed(2)} (from ${portions.totalImbueDiceCount}d${inputs.imbueDieType}) * (1 + (${powerForImbue} * ${imbueScaling * 100}%) / 100) = ${imbuePortion.toFixed(2)}`;
            this.scaledDiceScalingBreakdownSpan.innerHTML = portions.finalScaledDiceBreakdown || '0';
            this.imbuePowerSourceSpan.textContent = imbueUsesSpellpower ? `Spell Power (${spellPower})` : `Melee Power (${meleePower})`;
            this.scaledDiceAddedInputDisplaySpan.textContent = portions.totalAddedScaledDice;
            
            if (portions.addedDiceBreakdown && portions.addedDiceBreakdown.length > 0) {
                this.scaledDiceAddedInputDisplaySpan.textContent = portions.addedDiceBreakdown.join(' + ');
            } else {
                this.scaledDiceAddedInputDisplaySpan.textContent = '0';
            }

            this.reaperPenaltySpan.textContent = `${((1 - averages.reaperMultiplier) * 100).toFixed(1)}% Reduction`;

            // Show/hide the "+1" for the imbue toggle
            const showImbueToggleBonus = inputs.imbueDiceCount > 0;
            this.imbueToggleBonusSpan.classList.toggle('hidden', !showImbueToggleBonus);
            this.imbueToggleBonusSpan.previousElementSibling.classList.toggle('hidden', !showImbueToggleBonus); // Hides the '+' symbol

        }

        calculateDdoDamage() {
            // 1. Gather all data
            const inputs = this._getInputs();

            // 2. Calculate probabilities of outcomes
            const probabilities = this._calculateProbabilities(inputs);

            // 3. Calculate base damage for each component
            const portions = this._calculateDamagePortions(inputs);

            // 4. Calculate final average damage, applying probabilities and multipliers
            const averages = this._calculateAverages(portions, probabilities, inputs);

            // 5. Store results on the instance for the comparison table
            this.totalAvgBaseHitDmg = averages.base;
            this.totalAvgSneakDmg = averages.sneak;
            this.totalAvgImbueDmg = averages.imbue;
            this.totalAvgUnscaledDmg = averages.unscaled;
            this.totalAvgScaledDiceDmg = averages.scaledDice;

            // Calculate the grand total average damage
            this.totalAverageDamage = this.totalAvgBaseHitDmg + this.totalAvgSneakDmg + this.totalAvgImbueDmg + this.totalAvgUnscaledDmg + this.totalAvgScaledDiceDmg;

            // 6. Update the UI with all the calculated values
            this._updateSummaryUI(averages, portions, inputs);
            // 7. Update the per-roll breakdown table
            const critStartRoll = 21 - inputs.threatRange;
            const finalCritMult1920 = inputs.critMult + inputs.critMult1920;
            this.updateRollDamageTable({
                ...inputs,
                ...portions,
                ...averages,
                critStartRoll,
                finalCritMult1920,
                unscaled_normal_multi: inputs.unscaled.normal_multi,
                unscaled_normal_noMulti: inputs.unscaled.normal_noMulti,
                unscaled_crit_multi: inputs.unscaled.crit_multi,
                unscaled_crit_noMulti: inputs.unscaled.crit_noMulti
            });
        }

        /**
         * Populates the results table with damage for each d20 roll.
         */
        updateRollDamageTable(params) {
            const { missThreshold, grazeThreshold, critStartRoll, weaponPortion, seekerPortion, sneakPortion, imbuePortion, imbueCrits, grazePercent, critMult, finalCritMult1920, multiStrikeMultiplier, reaperMultiplier, unscaled_normal_multi, unscaled_normal_noMulti, unscaled_crit_multi, unscaled_crit_noMulti, scaledDice } = params;
            // Clear any existing rows from the table
            this.rollDamageTbody.innerHTML = '';

            for (let roll = 1; roll <= 20; roll++) {
                let baseDmg = 0, sneakDmg = 0, imbueDmg = 0;
                let rowClass = '';
                let outcome = '';

                if (roll <= missThreshold) { // Miss
                    baseDmg = 0;
                    sneakDmg = 0;
                    imbueDmg = 0;
                    outcome = 'Miss';
                } else if (roll <= grazeThreshold) { // Graze
                    baseDmg = weaponPortion * grazePercent;
                    sneakDmg = sneakPortion; // Sneak applies on graze
                    imbueDmg = 0; // Imbue does not apply on graze
                    rowClass = 'graze-row';
                    outcome = 'Graze';
                } else if (roll >= 19 && roll >= critStartRoll) { // Special 19-20 Critical Hit
                    baseDmg = (weaponPortion + seekerPortion) * finalCritMult1920;
                    sneakDmg = sneakPortion;
                    imbueDmg = imbuePortion * (1 + (imbueCrits ? 1 : 0));
                    rowClass = 'crit-row';
                    outcome = 'Critical';
                } else if (roll >= critStartRoll) { // Critical Hit
                    baseDmg = (weaponPortion + seekerPortion) * critMult;
                    sneakDmg = sneakPortion;
                    imbueDmg = imbuePortion * (1 + (imbueCrits ? 1 : 0));
                    rowClass = 'crit-row';
                    outcome = 'Critical';
                } else { // Normal Hit
                    baseDmg = weaponPortion;
                    sneakDmg = sneakPortion;
                    imbueDmg = imbuePortion;
                    outcome = 'Hit';
                }

                // Apply multi-strike multiplier to each component
                const finalBase = baseDmg * multiStrikeMultiplier * reaperMultiplier;
                const finalSneak = sneakDmg * multiStrikeMultiplier * reaperMultiplier;
                const finalImbue = imbueDmg * multiStrikeMultiplier * reaperMultiplier;

                let finalUnscaled = 0;
                if (outcome === 'Hit' || outcome === 'Critical') {
                    finalUnscaled += ((unscaled_normal_multi * multiStrikeMultiplier) + unscaled_normal_noMulti) * reaperMultiplier;
                }
                if (outcome === 'Critical') {
                    finalUnscaled += (((unscaled_crit_multi * multiStrikeMultiplier) + unscaled_crit_noMulti) * reaperMultiplier);
                }

                let finalScaledDice = 0;
                if (outcome !== 'Miss' && outcome !== 'Graze') {
                    finalScaledDice = scaledDice * multiStrikeMultiplier * reaperMultiplier;
                }
                const totalDamage = finalBase + finalSneak + finalImbue + finalUnscaled + finalScaledDice;

                // Create row and cells programmatically to prevent XSS.
                const row = document.createElement('tr');
                if (rowClass) {
                    row.className = rowClass;
                }

                const createCell = (text) => {
                    const cell = document.createElement('td');
                    cell.textContent = text;
                    return cell;
                };

                row.appendChild(createCell(roll));
                row.appendChild(createCell(finalBase.toFixed(2)));
                row.appendChild(createCell(finalSneak.toFixed(2)));
                row.appendChild(createCell(finalImbue.toFixed(2)));
                row.appendChild(createCell(finalUnscaled.toFixed(2)));
                row.appendChild(createCell(finalScaledDice.toFixed(2)));
                row.appendChild(createCell(totalDamage.toFixed(2)));
                row.appendChild(createCell(outcome));

                this.rollDamageTbody.appendChild(row);
            }
        }

        handleInputChange(action = null) {
            if (action) this.manager.recordAction(action);
            this.calculateDdoDamage();
            this.manager.updateComparisonTable();
            this.manager.saveState();
        }

        addEventListeners() {
            this.calculateBtn.addEventListener('click', () => this.calculateDdoDamage());

            // Add listeners for preset scaling buttons
            // These are simple and don't need complex removal logic if the whole set is destroyed.
            this.set75ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 75));
            this.set100ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 100));
            this.set150ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 150));
            this.set200ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 200));

            const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
            let oldValue = null;

            if (calculatorElement) {
                // Use 'mousedown' to capture the value right before a change might occur.
                // This works for clicks, typing, and stepper buttons.
                calculatorElement.addEventListener('mousedown', (e) => {
                    const targetTag = e.target.tagName;
                    if (targetTag === 'INPUT' || targetTag === 'SELECT') {
                        oldValue = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
                    }
                });

                // Use 'input' for text/number fields as it fires immediately.
                // Use 'change' for checkboxes.
                const recordChange = (e) => {
                    const targetTag = e.target.tagName;
                    if (targetTag === 'INPUT' || targetTag === 'SELECT') {
                        const newValue = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
                        // Only record if the value actually changed.
                        if (oldValue !== newValue) {
                            const action = {
                                type: 'VALUE_CHANGE',
                                setId: this.setId,
                                inputId: e.target.id,
                                oldValue: oldValue,
                                newValue: newValue
                            };
                            this.handleInputChange(action);
                            // CRITICAL FIX: Update oldValue after recording the change.
                            oldValue = newValue;
                        }
                    }
                };
                calculatorElement.addEventListener('input', recordChange);
                calculatorElement.addEventListener('change', recordChange);

                // Add a 'blur' event listener to the container to catch when a user leaves an input.
                // This will set empty number fields to 0 for a better user experience.
                calculatorElement.addEventListener('blur', (e) => {
                    const input = e.target;
                    // Check if the target is a number input and its value is empty.
                    const isNumericInput = input.type === 'number';
                    // Also check for text inputs that are used for damage notation.
                    const isDamageText = input.type === 'text' && (input.id.includes('unscaled-damage') || input.id.includes('weapon-damage'));

                    if (input.tagName === 'INPUT' && (isNumericInput || isDamageText) && input.value.trim() === '') {
                        input.value = '0';

                        // Manually trigger a 'change' event so the new '0' value is calculated and saved.
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }, true); // Use event capturing.
            }

            // Listener for adding a new unscaled damage row
            this.addUnscaledRowBtn.addEventListener('click', (e) => this.addUnscaledDamageRow(e));

            // Use event delegation for remove buttons. This single listener handles all
            // and for number input +/- buttons.
            // current and future remove buttons within this container.
            this.unscaledRowsContainer.addEventListener('click', (e) => {
                // Check if a remove button was clicked
                if (e.target && e.target.classList.contains('remove-row-btn')) {
                    e.preventDefault();
                    // Find the closest parent row and remove it
                    e.target.closest('.input-group-row').remove();
                    // Trigger a recalculation and save state
                    this.handleInputChange(); // This doesn't need an action, as removing rows isn't undoable (yet)
                }
            });

            // Listener for adding a new scaled dice damage row
            this.addScaledDiceRowBtn.addEventListener('click', (e) => this.addScaledDiceDamageRow(e));

            // Use event delegation for remove buttons within the scaled dice container
            this.scaledDiceRowsContainer.addEventListener('click', (e) => {
                if (e.target && e.target.classList.contains('remove-row-btn')) {
                    e.preventDefault();
                    e.target.closest('.input-group-row').remove();
                    this.handleInputChange();
                }
            });
        }
        removeEventListeners() { // Remove listeners from all inputs
            const calculatorElement = document.getElementById(`calculator-set-${this.setId}`);
            if (calculatorElement) {
                // Event listeners are on the container, which gets removed, so no need to manually remove them.
            }
            // With event delegation, we don't need to remove listeners from individual inputs.
            // The listeners are on the container, which gets removed from the DOM entirely
            // when a set is deleted, cleaning up the listeners automatically.
        }

        addUnscaledDamageRow(event) {
            event.preventDefault();
            const newIndex = Date.now(); // Use a timestamp for a unique ID to prevent collisions
            const idSuffix = this.idSuffix;
        
            // Clone the template content
            const template = document.getElementById('unscaled-row-template');
            if (!template) {
                console.error("Unscaled row template not found!");
                return;
            }
            const newRow = template.content.cloneNode(true).firstElementChild;
        
            // Find the next available number for the label
            const existingRows = this.unscaledRowsContainer.querySelectorAll('.input-group-row');
            const nextLabelNumber = existingRows.length + 1;
        
            // Update IDs and 'for' attributes to be unique
            newRow.querySelectorAll('[id*="-X"]').forEach(el => {
                el.id = el.id.replace('-X', `-${newIndex}${idSuffix}`);
            });
            newRow.querySelectorAll('[for*="-X"]').forEach(label => {
                label.setAttribute('for', label.getAttribute('for').replace('-X', `-${newIndex}${idSuffix}`));
            });
        
            // Update the main label text
            const mainLabel = newRow.querySelector('label[for^="unscaled-damage-"]');
            mainLabel.textContent = `Unscaled Damage ${nextLabelNumber}`;
        
            this.unscaledRowsContainer.appendChild(newRow);
        
            // Initialize adaptive sizing for the new inputs
            newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
        }

        addScaledDiceDamageRow(event) {
            event.preventDefault();
            const newIndex = Date.now(); // Use a timestamp for a unique ID
            const idSuffix = this.idSuffix;

            const newRow = document.createElement('div');
            newRow.className = 'input-group-row';
            newRow.innerHTML = `
                <input type="checkbox" id="scaled-dice-enabled-${newIndex}${idSuffix}" checked title="Enable/disable this scaled dice source.">
                <label for="scaled-dice-base-${newIndex}${idSuffix}">Base Dice</label>
                <input type="text" id="scaled-dice-base-${newIndex}${idSuffix}" value="1d6" class="adaptive-text-input">

                <label for="scaled-dice-proc-chance-${newIndex}${idSuffix}" class="short-label">Proc %</label>
                <input type="number" id="scaled-dice-proc-chance-${newIndex}${idSuffix}" value="100" min="0" max="100" class="small-input adaptive-text-input" title="Chance for this damage to occur on a hit">

                <label for="scaled-dice-scaling-percent-${newIndex}${idSuffix}" class="short-label">Scaling %</label>
                <input type="number" id="scaled-dice-scaling-percent-${newIndex}${idSuffix}" value="100" class="small-input" title="Percentage of Melee/Ranged Power to apply to this damage source.">
                
                <input type="checkbox" id="scaled-dice-scaling-toggle-${newIndex}${idSuffix}" checked>
                <label for="scaled-dice-scaling-toggle-${newIndex}${idSuffix}" class="inline-checkbox-label" title="Enable scaling with imbue dice count.">Scale with Imbue</label>
                
                <button class="remove-row-btn" title="Remove this damage source">&times;</button>
            `;

            this.scaledDiceRowsContainer.appendChild(newRow);

            // Manually trigger a change event so the new row is calculated and saved.
            const changeEvent = new Event('change', { bubbles: true });
            newRow.querySelector('input').dispatchEvent(changeEvent);

            // Initialize adaptive sizing for the new input
            newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));

            // Add a dummy action to the undo stack if needed, or just save state
            const action = {
                type: 'ADD_SCALED_ROW',
                setId: this.setId
            };
            this.handleInputChange(action);
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

            // Store unscaled damage rows
            state.unscaledDamageRows = [];
            this.unscaledRowsContainer.querySelectorAll('.input-group-row').forEach(row => {
                const dmgInput = row.querySelector(`input[id^="unscaled-damage-"]`);
                const procInput = row.querySelector(`input[id^="unscaled-proc-chance-"]`);
                const multiStrikeCheckbox = row.querySelector(`input[id^="unscaled-doublestrike-"]`);
                const onCritCheckbox = row.querySelector(`input[id^="unscaled-on-crit-"]`);

                if (dmgInput && procInput && multiStrikeCheckbox && onCritCheckbox) {
                    state.unscaledDamageRows.push({
                        damage: dmgInput.value,
                        procChance: procInput.value,
                        doublestrike: multiStrikeCheckbox.checked,
                        onCrit: onCritCheckbox.checked
                    });
                }
            });


            // Store scaled dice damage rows
            state.scaledDiceDamageRows = [];
            this.scaledDiceRowsContainer.querySelectorAll('.input-group-row').forEach(row => {
                const baseDiceInput = row.querySelector(`input[id^="scaled-dice-base-"]`);
                const scalingToggle = row.querySelector(`input[id^="scaled-dice-scaling-toggle-"]`);
                const scalingPercentInput = row.querySelector(`input[id^="scaled-dice-scaling-percent-"]`);
                const enabledCheckbox = row.querySelector(`input[id^="scaled-dice-enabled-"]`);
                const procChanceInput = row.querySelector(`input[id^="scaled-dice-proc-chance-"]`);
                if (baseDiceInput && scalingToggle && scalingPercentInput && enabledCheckbox && procChanceInput) {
                    state.scaledDiceDamageRows.push({
                        baseDice: baseDiceInput.value,
                        enableScaling: scalingToggle.checked,
                        procChance: procChanceInput.value,
                        scalingPercent: scalingPercentInput.value,
                        isEnabled: enabledCheckbox.checked
                    });
                }
            });

            return state;
        }

        // Retrieves the current name of the tab for this calculator instance
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

        setState(state) {
            if (!state) return;

            // Clear only the dynamic rows, not the entire container.
            // This preserves any static elements like headers or the "Dice Added" display.
            this.unscaledRowsContainer.querySelectorAll('.input-group-row').forEach(row => row.remove());
            this.scaledDiceRowsContainer.querySelectorAll('.input-group-row').forEach(row => {
                // Don't remove the read-only display box
                if (!row.classList.contains('read-only-display')) {
                    row.remove();
                }
            });

            const allInputs = document.querySelectorAll(`#calculator-set-${this.setId} input, #calculator-set-${this.setId} select`);
            allInputs.forEach(input => {
                const key = input.id.replace(`-set${this.setId}`, '');
                if (state.hasOwnProperty(key)) {
                    if (input.type === 'checkbox') {
                        input.checked = state[key];
                    } else {
                        input.value = state[key];
                    }
                }
            });

            // Recreate unscaled damage rows from state
            if (state.unscaledDamageRows) {
                state.unscaledDamageRows.forEach(rowData => {
                    this.addUnscaledDamageRow(new Event('dummy')); // Add a new empty row
                    const newRow = this.unscaledRowsContainer.lastElementChild;
                    // Find the inputs in the newly added row and set their values
                    const idSuffix = this.idSuffix;
                    newRow.querySelector(`input[id^="unscaled-damage-"]`).value = rowData.damage;
                    newRow.querySelector(`input[id^="unscaled-proc-chance-"]`).value = rowData.procChance;
                    newRow.querySelector(`input[id^="unscaled-doublestrike-"]`).checked = rowData.doublestrike;
                    newRow.querySelector(`input[id^="unscaled-on-crit-"]`).checked = rowData.onCrit;
                });
            } else {
                // If no unscaledDamageRows in state, but there were some default ones,
                // we need to initialize them based on how many were saved implicitly.
                // This handles cases where old saves without explicit rows still had default ones.
                const unscaledDamageKeys = Object.keys(state).filter(k => k.startsWith('unscaled-damage-'));
                let maxSavedUnscaledIdx = 0;
                unscaledDamageKeys.forEach(key => {
                    const idx = parseInt(key.match(/unscaled-damage-(\d+)/)[1], 10);
                    if (idx > maxSavedUnscaledIdx) maxSavedUnscaledIdx = idx;
                });

                // Add default unscaled rows if they exist in the state, up to maxSavedUnscaledIdx
                for (let i = 1; i <= maxSavedUnscaledIdx; i++) {
                    if (state.hasOwnProperty(`unscaled-damage-${i}`)) {
                        this.addUnscaledDamageRow(new Event('dummy'));
                        const newRow = this.unscaledRowsContainer.lastElementChild;
                        newRow.querySelector(`input[id^="unscaled-damage-"]`).value = state[`unscaled-damage-${i}`];
                        newRow.querySelector(`input[id^="unscaled-proc-chance-"]`).value = state[`unscaled-proc-chance-${i}`];
                        newRow.querySelector(`input[id^="unscaled-doublestrike-"]`).checked = state[`unscaled-doublestrike-${i}`];
                        newRow.querySelector(`input[id^="unscaled-on-crit-"]`).checked = state[`unscaled-on-crit-${i}`];
                    }
                }
            }


            // Recreate scaled dice damage rows from state
            if (state.scaledDiceDamageRows) {
                state.scaledDiceDamageRows.forEach(rowData => {
                    this.addScaledDiceDamageRow(new Event('dummy'));
                    const newRow = this.scaledDiceRowsContainer.lastElementChild;
                    newRow.querySelector(`input[id^="scaled-dice-base-"]`).value = rowData.baseDice;
                    newRow.querySelector(`input[id^="scaled-dice-scaling-toggle-"]`).checked = rowData.enableScaling;
                    newRow.querySelector(`input[id^="scaled-dice-enabled-"]`).checked = rowData.isEnabled;
                    newRow.querySelector(`input[id^="scaled-dice-proc-chance-"]`).value = rowData.procChance || 100;
                    newRow.querySelector(`input[id^="scaled-dice-scaling-percent-"]`).value = rowData.scalingPercent || 100;
                });
            }

            // After setting state, recalculate to update results
            this.calculateDdoDamage();

            // After all values are set, resize the inputs to fit the content
            this.resizeAllAdaptiveInputs();
        }

        applyValueChange(inputId, value) {
            const input = document.getElementById(inputId);
            if (!input) return;

            if (input.type === 'checkbox') {
                input.checked = value;
            } else {
                input.value = value;
            }
            // We don't record an action here because this is part of an undo/redo operation.
            // We just recalculate and save the new overall state.
            this.handleInputChange();
        }

        updateSummaryHeader() {
            const tabName = this.getTabName();
            if (this.summaryHeader) {
                this.summaryHeader.textContent = `Summary of ${tabName}`;
            }
        }

        handleSetScalingClick(e, value) {
            e.preventDefault();
            this.imbueScalingInput.value = value;
            // Manually trigger a change event so the calculation updates and state is saved
            const changeEvent = new Event('change', { bubbles: true });
            this.imbueScalingInput.dispatchEvent(changeEvent);
        }
    }