import { parseDiceNotation } from './utils.js';
import { BaseCalculator } from './BaseCalculator.js';

const defaultState = {
    'weapon-dice': 7,
    'weapon-damage': '1d10+3',
    'bonus-base-damage': 125,
    'crit-threat': '15-20',
    'crit-multiplier': 4,
    'seeker-damage': '0',
    'crit-multiplier-19-20': 2,
    'miss-threshold': 1,
    'graze-threshold': 5,
    'graze-percent': 20,
    'doublestrike': 0,
    'is-doubleshot': false,
    'melee-power': 0,
    'spell-power': 0,
    'archers-focus': 0,
    'improved-archers-focus': false,
    'reaper-skulls': '0',
    'sneak-attack-dice': 0,
    'sneak-bonus': 0,
    'imbue-active': false,
    'imbue-dice-count': 0,
    'imbue-die-type': 6,
    'imbue-scaling': 100,
    'imbue-uses-spellpower': false,
    'imbue-crits': false,
    unscaledRows: [],
    scaledDiceRows: [],
};

export class WeaponCalculator extends BaseCalculator {
        constructor(setId, manager, name) { // Add 'name' to the constructor
            super(setId, manager, name);
            // Properties to store calculation results for the comparison table
            this.state = JSON.parse(JSON.stringify(defaultState)); // Deep copy

            this.totalAverageDamage = 0;
            this.totalAvgBaseHitDmg = 0;
            this.totalAvgSneakDmg = 0;
            this.totalAvgUnscaledDmg = 0; // New property for unscaled damage
            this.totalAvgImbueDmg = 0; // This will hold the sum of all unscaled sources
            this.totalAvgScaledDiceDmg = 0;

            this.getElements();
            this.addEventListeners();
            // We will call the initial calculation from the manager

            // Initialize adaptive sizing for all relevant inputs
            this._initializeAdaptiveInputs();
        }

        getElements() {
            const get = (elementName) => this.container.querySelector(`[data-element="${elementName}"]`);

            this.weaponDiceInput = get('weapon-dice');
            this.weaponDamageInput = get('weapon-damage');
            this.bonusBaseDamageInput = get('bonus-base-damage');
            this.meleePowerInput = get('melee-power');
            this.spellPowerInput = get('spell-power');
            this.rangedPowerBonusSpan = get('ranged-power-bonus');
            this.archersFocusSlider = get('archers-focus');
            this.archersFocusValueSpan = get('archers-focus-value');
            this.improvedArchersFocusCheckbox = get('improved-archers-focus');
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
            this.imbueActiveCheckbox = get('imbue-active');
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
            this.avgUnscaledDamageSpan = get('avg-unscaled-damage');
            this.avgScaledDiceDamageSpan = get('avg-scaled-dice-damage');
            this.totalAvgDamageSpan = get('total-avg-damage');
            this.weaponScalingSpan = get('weapon-scaling');
            this.sneakScalingSpan = get('sneak-scaling');
            this.imbueScalingBreakdownSpan = get('imbue-scaling-breakdown');
            this.unscaledScalingBreakdownSpan = get('unscaled-scaling-breakdown');
            this.scaledDiceScalingBreakdownSpan = get('scaled-dice-scaling-breakdown');
            this.scaledDiceAddedInputDisplaySpan = get('scaled-dice-added-input-display');
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
        _mapStateToInputs() {
            const isDoubleshot = this.state['is-doubleshot'];
            let multiStrikeValue = parseFloat(this.state['doublestrike']) || 0;
            if (!isDoubleshot) {
                multiStrikeValue = Math.min(multiStrikeValue, 100);
            }
        
            const unscaled = {
                normal_multi: 0,
                normal_noMulti: 0,
                crit_multi: 0,
                crit_noMulti: 0
            };
        
            let totalUnscaledAverage = 0;
            
            // Process unscaled rows from state
            for (const key in this.state) {
                if (key.startsWith('unscaled-damage-')) {
                    const id = key.substring('unscaled-damage-'.length);
                    const damage = parseDiceNotation(this.state[key]);
                    const procChance = (parseFloat(this.state[`unscaled-proc-chance-${id}`]) || 100) / 100;
                    const averageDamage = damage * procChance;
                    totalUnscaledAverage += averageDamage;
        
                    const multiStrike = this.state[`unscaled-doublestrike-${id}`];
                    const onCrit = this.state[`unscaled-on-crit-${id}`];
        
                    if (multiStrike) {
                        if (onCrit) unscaled.crit_multi += averageDamage;
                        else unscaled.normal_multi += averageDamage;
                    } else {
                        if (onCrit) unscaled.crit_noMulti += averageDamage;
                        else unscaled.normal_noMulti += averageDamage;
                    }
                }
            }
        
            const scaledDiceDamage = [];
            // Process scaled dice rows from state
            for (const key in this.state) {
                if (key.startsWith('scaled-dice-enabled-') && this.state[key]) {
                    const id = key.substring('scaled-dice-enabled-'.length);
                    scaledDiceDamage.push({
                        baseDice: this.state[`scaled-dice-base-${id}`],
                        procChance: (parseFloat(this.state[`scaled-dice-proc-chance-${id}`]) || 100) / 100,
                        enableScaling: this.state[`scaled-dice-scaling-toggle-${id}`],
                        scalingPercent: parseFloat(this.state[`scaled-dice-scaling-percent-${id}`]) || 100,
                        isEnabled: true
                    });
                }
            }
        
            return {
                additionalWeaponDice: parseFloat(this.state['weapon-dice']) || 0,
                parsedWeaponDmg: parseDiceNotation(this.state['weapon-damage']) || 0,
                bonusBaseDmg: parseFloat(this.state['bonus-base-damage']) || 0,
                // Add Archer's Focus bonus to the base Melee/Ranged Power
                meleePower: (parseFloat(this.state['melee-power']) || 0) + ((parseInt(this.state['archers-focus'], 10) || 0) * 5),
                spellPower: parseFloat(this.state['spell-power']) || 0,
                archersFocus: parseInt(this.state['archers-focus'], 10) || 0,
                threatRange: this.parseThreatRange(this.state['crit-threat']),
                critMult: parseFloat(this.state['crit-multiplier']) || 2,
                critMult1920: parseFloat(this.state['crit-multiplier-19-20']) || 0,
                seekerDmg: parseDiceNotation(this.state['seeker-damage']),
                sneakDiceCount: parseInt(this.state['sneak-attack-dice']) || 0,
                sneakBonusDmg: parseFloat(this.state['sneak-bonus']) || 0,
                missThreshold: Math.max(1, parseInt(this.state['miss-threshold']) || 1),
                grazeThreshold: parseInt(this.state['graze-threshold']) || 0,
                grazePercent: (parseFloat(this.state['graze-percent']) || 0) / 100,
                reaperSkulls: parseInt(this.state['reaper-skulls']) || 0,
                isImbueActive: this.state['imbue-active'],
                imbueDiceCount: parseInt(this.state['imbue-dice-count']) || 0,
                imbueDieType: parseInt(this.state['imbue-die-type']) || 6,
                imbueScaling: (parseFloat(this.state['imbue-scaling']) || 100) / 100,
                imbueCrits: this.state['imbue-crits'],
                imbueUsesSpellpower: this.state['imbue-uses-spellpower'],
                doublestrikeChance: multiStrikeValue / 100,
                isDoubleshot: isDoubleshot,
                unscaled: unscaled,
                scaledDiceDamage: scaledDiceDamage,
                totalUnscaledAverage: totalUnscaledAverage
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
                seekerDmg, sneakDiceCount, sneakBonusDmg, isImbueActive, imbueDiceCount, imbueDieType,
                imbueScaling, imbueUsesSpellpower, scaledDiceDamage
            } = inputs;

            const baseDmg = (parsedWeaponDmg * additionalWeaponDice) + bonusBaseDmg;
            const powerMultiplier = 1 + (meleePower / 100);
            const weaponPortion = baseDmg * powerMultiplier;
            const seekerPortion = seekerDmg * powerMultiplier;

            const sneakDiceDmg = sneakDiceCount * 3.5;
            const sneakPortion = (sneakDiceDmg + sneakBonusDmg) * (1 + (meleePower * 1.5) / 100);

            const totalImbueDiceCount = isImbueActive ? imbueDiceCount + 1 : 0;
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
                let breakdown = `(Base (${baseDiceAvg.toFixed(2)}) + Imbue (${additionalDiceAvg.toFixed(2)})) &times; Power (${powerMultiplier.toFixed(2)}) &times; Proc (${scaledDmg.procChance.toFixed(2)}) = ${lineResult.toFixed(2)}`;
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

            //weapon breakdown
            const weaponDiceAvg = (inputs.parsedWeaponDmg * inputs.additionalWeaponDice);
            const weaponPowerMod = (1 + inputs.meleePower / 100);
            const weaponAfterPowerMod = portions.weaponPortion;
            const multiStrike = averages.multiStrikeMultiplier;
            this.weaponScalingSpan.innerHTML = `${weaponDiceAvg.toFixed(2)} + ${inputs.bonusBaseDmg.toFixed(2)} &times; Power Mod (${weaponPowerMod.toFixed(2)}) = ${weaponAfterPowerMod.toFixed(2)} &times; DS (${multiStrike.toFixed(2)}) = <strong>${(weaponAfterPowerMod * multiStrike).toFixed(2)}</strong>`;

            //sneak breakdown   
            const sneakPowerMod = (1 + (meleePower * 1.5) / 100);
            const sneakAfterPowerMod = sneakPortion;
            this.sneakScalingSpan.innerHTML = `Dice Avg (${sneakDiceDmg.toFixed(2)}) + Flat (${sneakBonusDmg.toFixed(2)}) &times; Power Mod (${sneakPowerMod.toFixed(2)}) = ${sneakAfterPowerMod.toFixed(2)} &times; DS (${multiStrike.toFixed(2)}) = <strong>${(sneakAfterPowerMod * multiStrike).toFixed(2)}</strong>`;
            
            //imbue breakdown
            const imbuePowerMod = (1 + (powerForImbue * imbueScaling) / 100);
            const imbueAfterPowerMod = imbuePortion;
            this.imbueScalingBreakdownSpan.innerHTML = `Dice Avg (${portions.imbueDice.toFixed(2)}) &times; Power Mod (${imbuePowerMod.toFixed(2)}) = ${imbueAfterPowerMod.toFixed(2)} &times; DS (${multiStrike.toFixed(2)}) = <strong>${(imbueAfterPowerMod * multiStrike).toFixed(2)}</strong>`;
            
            //unscaled breakdown
            const unscaledWithMulti = inputs.unscaled.normal_multi + inputs.unscaled.crit_multi;
            const unscaledWithoutMulti = inputs.unscaled.normal_noMulti + inputs.unscaled.crit_noMulti;
            const unscaledBreakdownText = `
                No DS (${unscaledWithoutMulti.toFixed(2)})    
                || DS (${unscaledWithMulti.toFixed(2)}) &times; Multiplier (${multiStrike.toFixed(2)})
                = <strong>${(unscaledWithMulti * averages.multiStrikeMultiplier + unscaledWithoutMulti).toFixed(2)}</strong>
            `;
            this.unscaledScalingBreakdownSpan.innerHTML = unscaledBreakdownText;

            //scaled dice breakdown
            const totalScaledDiceAfterPowerMod = portions.totalAvgScaledDiceDmg;
            this.scaledDiceScalingBreakdownSpan.innerHTML = portions.finalScaledDiceBreakdown ? `${portions.finalScaledDiceBreakdown} &times; DS (${multiStrike.toFixed(2)}) = <strong>${(totalScaledDiceAfterPowerMod * multiStrike).toFixed(2)}</strong>` : '0';
            this.imbuePowerSourceSpan.textContent = imbueUsesSpellpower ? `Spell Power (${spellPower})` : `Melee Power (${meleePower})`;
            this.scaledDiceAddedInputDisplaySpan.textContent = portions.totalAddedScaledDice;
            
            if (portions.addedDiceBreakdown && portions.addedDiceBreakdown.length > 0) {
                this.scaledDiceAddedInputDisplaySpan.textContent = portions.addedDiceBreakdown.join(' + ');
            } else {
                this.scaledDiceAddedInputDisplaySpan.textContent = '0';
            }

            this.reaperPenaltySpan.textContent = `${((1 - averages.reaperMultiplier) * 100).toFixed(1)}% Reduction`;

            // Show/hide the "+1" for the imbue toggle
            const showImbueToggleBonus = inputs.isImbueActive;
            this.imbueToggleBonusSpan.classList.toggle('hidden', !showImbueToggleBonus);
            this.imbueToggleBonusSpan.previousElementSibling.classList.toggle('hidden', !showImbueToggleBonus); // Hides the '+' symbol

            // Update Archer's Focus display
            const archersFocusValue = inputs.archersFocus;
            const archersFocusBonus = archersFocusValue * 5;
            this.archersFocusValueSpan.textContent = archersFocusValue;
            this.rangedPowerBonusSpan.textContent = archersFocusBonus;

            // Update slider max based on the "Improved" checkbox
            const isImproved = this.state['improved-archers-focus'];
            this.archersFocusSlider.max = isImproved ? '25' : '15';

            // If the slider's value is now greater than its max, clamp it.
            // This handles unchecking "Improved" when the value is > 15.
            if (parseInt(this.archersFocusSlider.value, 10) > parseInt(this.archersFocusSlider.max, 10)) {
                this.archersFocusSlider.value = this.archersFocusSlider.max;
                this.state['archers-focus'] = this.archersFocusSlider.value; // Also update the state directly
            }
        }

        calculateDdoDamage() {
            // 1. Gather all data
            const inputs = this._mapStateToInputs();

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

            this.calculateDdoDamage();
            this.manager.updateComparisonTable();
            this.manager.saveState();

            // If the "Improved" checkbox was just changed, we must check if the slider value is now invalid.
            if (key === 'improved-archers-focus') {
                const max = this.archersFocusSlider.max; // The max has already been updated by the UI render
                if (parseInt(this.state['archers-focus'], 10) > parseInt(max, 10)) {
                    // If the value is out of bounds, update the state and trigger a new input event to recalculate everything.
                    this.state['archers-focus'] = max;
                    this.archersFocusSlider.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }
        }

        setState(state) {
            // Perform a deep merge of the new state into the existing state
            this.state = { ...this.state, ...state };
            super.setState(this.state); // This calls deserializeForm
            this.calculateDdoDamage();
        }

        addEventListeners() {
            super.addEventListeners();
            this.calculateBtn.addEventListener('click', () => this.calculateDdoDamage());

            // Add listeners for preset scaling buttons
            // These are simple and don't need complex removal logic if the whole set is destroyed.
            this.set75ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 75));
            this.set100ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 100));
            this.set150ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 150));
            this.set200ScalingBtn?.addEventListener('click', (e) => this.handleSetScalingClick(e, 200));

            if (this.container) {
                const boundHandler = this.handleInputChange.bind(this);
                this.container.addEventListener('input', boundHandler);
                this.container.addEventListener('change', boundHandler);

                this.container.addEventListener('blur', (e) => {
                    const input = e.target;
                    const isNumericInput = input.type === 'number';
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
                const removeBtn = e.target.closest('.remove-row-btn');
                if (removeBtn) {
                    e.preventDefault();
                    const row = removeBtn.closest('.input-group-row');
                    if (!row) return;

                    const rowId = row.dataset.rowId;
                    const rowData = {};
                    row.querySelectorAll('input[data-element]').forEach(input => {
                        rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
                    });

                    const parent = row.parentNode;
                    const rowIndex = Array.prototype.indexOf.call(parent.children, row);

                    this.manager.recordAction({ type: 'REMOVE_DYNAMIC_ROW', setId: this.setId, rowType: 'unscaled', rowId, rowData, rowIndex });

                    row.remove();
                    this.calculateDdoDamage();
                    this.manager.saveState();
                }
            });

            // Listener for adding a new scaled dice damage row
            this.addScaledDiceRowBtn.addEventListener('click', (e) => this.addScaledDiceDamageRow(e));

            // Use event delegation for remove buttons within the scaled dice container
            this.scaledDiceRowsContainer.addEventListener('click', (e) => {
                const removeBtn = e.target.closest('.remove-row-btn');
                if (removeBtn) {
                    e.preventDefault();
                    const row = removeBtn.closest('.input-group-row');
                    if (!row) return;

                    const rowId = row.dataset.rowId;
                    const rowData = {};
                    row.querySelectorAll('input[data-element]').forEach(input => {
                        rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
                    });

                    const parent = row.parentNode;
                    const rowIndex = Array.prototype.indexOf.call(parent.children, row);

                    this.manager.recordAction({ type: 'REMOVE_DYNAMIC_ROW', setId: this.setId, rowType: 'scaled', rowId, rowData, rowIndex });

                    row.remove();
                    this.calculateDdoDamage();
                    this.manager.saveState();
                }
            });
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
            newRow.dataset.rowId = newIndex;
        
            // Find the next available number for the label
            const existingRows = this.unscaledRowsContainer.querySelectorAll('.input-group-row');
            const nextLabelNumber = existingRows.length + 1;
        
            // Update IDs and 'for' attributes to be unique
            newRow.querySelectorAll('[id*="-X"]').forEach(el => {
                el.dataset.element = el.dataset.element.replace('-X', `-${newIndex}`);
                el.id = el.id.replace('-X', `-${newIndex}${idSuffix}`);
            });
            newRow.querySelectorAll('[for*="-X"]').forEach(label => {
                const oldFor = label.getAttribute('for');
                label.setAttribute('for', oldFor.replace('-X', `-${newIndex}${idSuffix}`));
            });
        
            // Update the main label text
            const mainLabel = newRow.querySelector('label[for^="unscaled-damage-"]');
            mainLabel.textContent = `Unscaled Damage ${nextLabelNumber}`;
        
            this.unscaledRowsContainer.appendChild(newRow);
        
            // Initialize adaptive sizing for the new inputs
            newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
        
            const rowData = {};
            newRow.querySelectorAll('input[data-element]').forEach(input => {
                rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
            });
            const parent = newRow.parentNode;
            const rowIndex = Array.prototype.indexOf.call(parent.children, newRow);

            this.manager.recordAction({ type: 'ADD_DYNAMIC_ROW', setId: this.setId, rowType: 'unscaled', rowId: newIndex, rowData, rowIndex });
            this.calculateDdoDamage();
            this.manager.saveState();
        }

        addScaledDiceDamageRow(event) {
            event.preventDefault();
            const newIndex = Date.now(); // Use a timestamp for a unique ID
            const idSuffix = this.idSuffix;

            const newRow = document.createElement('div');
            newRow.className = 'input-group-row';
            newRow.dataset.rowId = newIndex;
            newRow.innerHTML = `
                <input type="checkbox" data-element="scaled-dice-enabled-${newIndex}" id="scaled-dice-enabled-${newIndex}${idSuffix}" checked title="Enable/disable this scaled dice source.">
                <label for="scaled-dice-base-${newIndex}${idSuffix}" >Base Dice</label>
                <input type="text" data-element="scaled-dice-base-${newIndex}" id="scaled-dice-base-${newIndex}${idSuffix}" value="1d6" class="adaptive-text-input">

                <label for="scaled-dice-proc-chance-${newIndex}${idSuffix}" class="short-label" >Proc %</label>
                <input type="number" data-element="scaled-dice-proc-chance-${newIndex}" id="scaled-dice-proc-chance-${newIndex}${idSuffix}" value="100" min="0" max="100" class="small-input adaptive-text-input" title="Chance for this damage to occur on a hit">

                <label for="scaled-dice-scaling-percent-${newIndex}${idSuffix}" class="short-label" >Scaling %</label>
                <input type="number" data-element="scaled-dice-scaling-percent-${newIndex}" id="scaled-dice-scaling-percent-${newIndex}${idSuffix}" value="100" class="small-input" title="Percentage of Melee/Ranged Power to apply to this damage source.">
                
                <input type="checkbox" data-element="scaled-dice-scaling-toggle-${newIndex}" id="scaled-dice-scaling-toggle-${newIndex}${idSuffix}" checked>
                <label for="scaled-dice-scaling-toggle-${newIndex}${idSuffix}" class="inline-checkbox-label"  title="Enable scaling with imbue dice count.">Scale with Imbue</label>
                
                <button class="remove-row-btn"  title="Remove this damage source">&times;</button>
            `;

            this.scaledDiceRowsContainer.appendChild(newRow);

            // Manually trigger a change event so the new row is calculated and saved.
            const changeEvent = new Event('change', { bubbles: true });
            newRow.querySelector('input').dispatchEvent(changeEvent);

            // Initialize adaptive sizing for the new input
            newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));

            const rowData = {};
            newRow.querySelectorAll('input[data-element]').forEach(input => {
                rowData[input.dataset.element] = input.type === 'checkbox' ? input.checked : input.value;
            });
            const parent = newRow.parentNode;
            const rowIndex = Array.prototype.indexOf.call(parent.children, newRow);

            this.manager.recordAction({ type: 'ADD_DYNAMIC_ROW', setId: this.setId, rowType: 'scaled', rowId: newIndex, rowData, rowIndex });
            this.calculateDdoDamage();
            this.manager.saveState();
        }

        removeDynamicRow(rowId, rowType) {
            const container = rowType === 'unscaled' ? this.unscaledRowsContainer : this.scaledDiceRowsContainer;
            const row = container.querySelector(`[data-row-id="${rowId}"]`);
            if (row) {
                row.remove();
                this.calculateDdoDamage();
            }
        }

        recreateDynamicRow(rowId, rowType, rowData, rowIndex) {
            const idSuffix = this.idSuffix;
            let newRow;
            let container;

            if (rowType === 'unscaled') {
                container = this.unscaledRowsContainer;
                const template = document.getElementById('unscaled-row-template');
                newRow = template.content.cloneNode(true).firstElementChild;
                newRow.dataset.rowId = rowId;

                const nextLabelNumber = container.querySelectorAll('.input-group-row').length + 1;
                newRow.querySelector('label[for^="unscaled-damage-"]').textContent = `Unscaled Damage ${nextLabelNumber}`;

                newRow.querySelectorAll('[id*="-X"]').forEach(el => {
                    el.id = el.id.replace('-X', `-${rowId}${idSuffix}`);
                });
                newRow.querySelectorAll('[for*="-X"]').forEach(label => {
                    label.setAttribute('for', label.getAttribute('for').replace('-X', `-${rowId}${idSuffix}`));
                });
            } else { // scaled
                container = this.scaledDiceRowsContainer;
                newRow = document.createElement('div');
                newRow.className = 'input-group-row';
                newRow.dataset.rowId = rowId;
                newRow.innerHTML = `
                    <input type="checkbox" data-element="scaled-dice-enabled-${rowId}" id="scaled-dice-enabled-${rowId}${idSuffix}" title="Enable/disable this scaled dice source.">
                    <label for="scaled-dice-base-${rowId}${idSuffix}">Base Dice</label>
                    <input type="text" data-element="scaled-dice-base-${rowId}" id="scaled-dice-base-${rowId}${idSuffix}" class="adaptive-text-input">
                    <label for="scaled-dice-proc-chance-${rowId}${idSuffix}" class="short-label">Proc %</label>
                    <input type="number" data-element="scaled-dice-proc-chance-${rowId}" id="scaled-dice-proc-chance-${rowId}${idSuffix}" class="small-input adaptive-text-input" title="Chance for this damage to occur on a hit">
                    <label for="scaled-dice-scaling-percent-${rowId}${idSuffix}" class="short-label">Scaling %</label>
                    <input type="number" data-element="scaled-dice-scaling-percent-${rowId}" id="scaled-dice-scaling-percent-${rowId}${idSuffix}" class="small-input" title="Percentage of Melee/Ranged Power to apply to this damage source.">
                    <input type="checkbox" data-element="scaled-dice-scaling-toggle-${rowId}" id="scaled-dice-scaling-toggle-${rowId}${idSuffix}">
                    <label for="scaled-dice-scaling-toggle-${rowId}${idSuffix}" class="inline-checkbox-label" title="Enable scaling with imbue dice count.">Scale with Imbue</label>
                    <button class="remove-row-btn" title="Remove this damage source">&times;</button>
                `;
            }

            newRow.querySelectorAll('input[data-element]').forEach(input => {
                const key = input.dataset.element;
                if (rowData.hasOwnProperty(key)) {
                    input.type === 'checkbox' ? (input.checked = rowData[key]) : (input.value = rowData[key]);
                }
            });

            container.insertBefore(newRow, container.children[rowIndex]);
            newRow.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
            this.calculateDdoDamage();
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