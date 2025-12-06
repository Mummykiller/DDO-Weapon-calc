import { WeaponCalculator } from './WeaponCalculator.js';
import { SpellCalculator } from './SpellCalculator.js';

// Wait for the entire page to load before running the script
document.addEventListener('DOMContentLoaded', () => {
    class CalculatorManager {
        constructor() {
            this.calculators = new Map();
            this.navContainer = document.querySelector('.set-navigation');
            this.setsContainer = document.querySelector('.calculator-wrapper');
            this.addSetBtn = document.getElementById('add-weapon-set-btn');
            this.addSpellSetBtn = document.getElementById('add-spell-set-btn');
            this.comparisonTbody = document.getElementById('comparison-tbody'); // This is fine
            this.templateHTML = this.getTemplateHTML();

            // Import/Export Modal Elements
            this.importBtn = document.getElementById('import-btn');
            this.exportBtn = document.getElementById('export-btn');
            this.modalBackdrop = document.getElementById('modal-backdrop');
            this.modalTitle = document.getElementById('modal-title');
            this.modalDescription = document.getElementById('modal-description');
            this.modalTextarea = document.getElementById('modal-textarea');
            this.modalCopyBtn = document.getElementById('modal-copy-btn');
            this.modalSaveFileBtn = document.getElementById('modal-save-file-btn');
            this.modalLoadBtn = document.getElementById('modal-load-btn');
            this.modalFileInput = document.getElementById('modal-file-input');
            this.formatJsonBtn = document.getElementById('format-json-btn');
            this.formatSummaryBtn = document.getElementById('format-summary-btn');
            this.modalCloseBtn = document.getElementById('modal-close-btn');

            // Undo/Redo stacks
            this.undoStack = [];
            this.redoStack = [];

            this.activeEnterHandler = null;
            this.activeContainerForEnter = null;

            this.activeSetId = 1;
            this.nextSetId = 1; // Moved nextSetId into CalculatorManager

            this.addSetBtn.addEventListener('click', () => this.addNewSet('weapon'));
            this.addSpellSetBtn.addEventListener('click', () => this.addNewSet('spell'));
            // Try to load state. If it fails (e.g., first visit), create the initial set.
            if (!this.loadState()) {
                this.addNewSet('weapon', 1);
                this.calculators.get(1)?.calculateDdoDamage(); // Perform initial calculation
            }

            this.addDragAndDropListeners();
            this.addImportExportListeners();
            this.addUndoRedoListeners();
        }


        addNewSet(type = 'weapon', setIdToUse = null, index = -1) {
            if (this.calculators.size >= 6) {
                alert("You have reached the maximum of 6 sets.");
                return;
            }

            const isSpell = type === 'spell';
            const activeCalc = this.calculators.get(this.activeSetId);
            const stateToCopy = (setIdToUse === null && (
                (isSpell && activeCalc instanceof SpellCalculator) || 
                (!isSpell && activeCalc instanceof WeaponCalculator)
            )) ? activeCalc.getState() : null;

            let newSetId;
            if (setIdToUse !== null) {
                newSetId = setIdToUse;
                if (newSetId >= this.nextSetId) {
                    this.nextSetId = newSetId + 1;
                }
            } else {
                newSetId = this.findNextAvailableId();
            }

            const templateId = isSpell ? 'spell-calculator-template' : 'calculator-set-template';
            const templateNode = document.getElementById(templateId).content.cloneNode(true);

            let modifiedInnerHtml = templateNode.firstElementChild.outerHTML.replace(/\s(id)="([^"]+)"/g, (match, attr, id) => {
                return ` id="${id}-set${newSetId}"`;
            });
            modifiedInnerHtml = modifiedInnerHtml.replace(/for="([^"]+)"/g, (match, id) => {
                return `for="${id}-set${newSetId}"`;
            });

            const newSetContainer = document.createElement('div');
            newSetContainer.id = `calculator-set-${newSetId}`;
            newSetContainer.className = 'calculator-container calculator-set';
            newSetContainer.innerHTML = modifiedInnerHtml;

            const allContainers = this.setsContainer.querySelectorAll('.calculator-set');
            if (index !== -1 && index < allContainers.length) {
                this.setsContainer.insertBefore(newSetContainer, allContainers[index]);
            } else {
                this.setsContainer.appendChild(newSetContainer);
            }

            const tab = this.createTab(newSetId);
            if (isSpell) {
                tab.classList.add('spell-tab-indicator');
            }
            const allTabs = this.navContainer.querySelectorAll('.nav-tab');
            if (index !== -1 && index < allTabs.length) {
                this.navContainer.insertBefore(tab, allTabs[index]);
            } else {
                const firstButton = this.navContainer.querySelector('.nav-action-btn');
                this.navContainer.insertBefore(tab, firstButton || null);
            }

            const CalculatorClass = isSpell ? SpellCalculator : WeaponCalculator;
            const setName = isSpell ? `Spell Set ${newSetId}` : `Set ${newSetId}`;
            this.calculators.set(newSetId, new CalculatorClass(newSetId, this, setName));
            
            const newCalc = this.calculators.get(newSetId);

            if (stateToCopy && newCalc) {
                newCalc.setState(stateToCopy);
            }

            this.switchToSet(newSetId);
            if (!this.isLoading) {
                this.saveState();
            }
            this.updateComparisonTable();
        }

        createTab(setId) {
            const tab = document.createElement('div');
            tab.className = 'nav-tab';
            tab.draggable = true; // Make the tab draggable
            tab.dataset.set = setId;

            const tabNameSpan = document.createElement('span');
            tabNameSpan.className = 'tab-name';
            tabNameSpan.textContent = `Set ${setId}`;
            tabNameSpan.contentEditable = true;
            tabNameSpan.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    tabNameSpan.blur();
                }
            });

            // We'll handle recording the name change via focus/blur on the span
            this.addNameChangeListeners(tabNameSpan, setId);

            const closeBtn = document.createElement('button');
            closeBtn.className = 'close-tab-btn';
            closeBtn.innerHTML = '&times;';
            closeBtn.title = 'Remove this set';
            closeBtn.addEventListener('click', (e) => {
                e.stopPropagation(); // Prevent tab switch when closing
                this.removeSet(setId);
            });

            tab.appendChild(tabNameSpan);
            tab.appendChild(closeBtn);
            tab.addEventListener('click', () => this.switchToSet(setId));
            return tab;
        }

        addNameChangeListeners(span, setId) {
            let oldName = '';
            span.addEventListener('focus', () => {
                oldName = span.textContent;
            });
            span.addEventListener('blur', () => {
                const newName = span.textContent;
                if (oldName !== newName) {
                    this.recordAction({
                        type: 'RENAME',
                        setId: setId,
                        oldValue: oldName,
                        newValue: newName
                    });
                    // Trigger a save and update
                    const calc = this.calculators.get(setId);
                    if (calc) {
                        if (calc instanceof WeaponCalculator) {
                            calc.handleInputChange();
                        } else {
                            calc.handleInputChange({ target: span }); // Simulate a change
                        }
                        if (this.activeSetId === setId) calc.updateSummaryHeader();
                    }
                }
            });
        }

        recreateSet(setId, state, index) {
            this.addNewSet(state.type, setId, index);
            const newCalc = this.calculators.get(setId);
            if (newCalc) {
                newCalc.setState(state);
                newCalc.setTabName(state.tabName);
                if (newCalc instanceof WeaponCalculator) {
                    newCalc.updateSummaryHeader();
                }
            }
            // When undoing a tab closure, we don't want to automatically switch to it.
            // We'll just make sure the currently active tab remains visibly active.
            // The new tab will be created but won't be active unless it was the only one.
            this.switchToSet(this.activeSetId || setId);
        }

        removeSet(setId, isLoading = false) {
            if (this.calculators.size <= 1) {
                alert("You cannot remove the last set.");
                return;
            }

            // Clean up
            const calcToRemove = this.calculators.get(setId);

            // If this is a user action (not part of loading or undo), record it.
            if (!isLoading) {
                const state = calcToRemove.getState();
                state.tabName = calcToRemove.getTabName();
                state.type = calcToRemove instanceof SpellCalculator ? 'spell' : 'weapon';
                const tabs = [...this.navContainer.querySelectorAll('.nav-tab')];
                const index = tabs.findIndex(tab => parseInt(tab.dataset.set, 10) === setId);

                this.recordAction({ type: 'REMOVE_SET', setId, state, index });
            }

            if (calcToRemove && typeof calcToRemove.removeEventListeners === 'function') {
                calcToRemove.removeEventListeners();
            }
            this.calculators.delete(setId);

            document.getElementById(`calculator-set-${setId}`).remove();
            document.querySelector(`.nav-tab[data-set="${setId}"]`).remove();

            // If we deleted the active tab, switch to a new one
            if (this.activeSetId === setId) {
                if (this.calculators.size > 0) {
                    const firstSetId = this.calculators.keys().next().value;
                    this.switchToSet(firstSetId);
                } else {
                    this.activeSetId = null; // No active set
                }
            }
            // Otherwise, the active tab remains the same, which is fine.

            if (!isLoading) {
                this.saveState();
                this.updateComparisonTable();
            }
        }

        switchToSet(setId) {
            // If there's an old handler on a previous container, remove it first.
            if (this.activeContainerForEnter && this.activeEnterHandler) {
                this.activeContainerForEnter.removeEventListener('keydown', this.activeEnterHandler);
            }

            // Deactivate all
            document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.calculator-set').forEach(s => s.classList.remove('active'));

            // Activate the selected one
            const newTab = document.querySelector(`.nav-tab[data-set="${setId}"]`);
            const newContainer = document.getElementById(`calculator-set-${setId}`);

            if (!newTab || !newContainer) return; // Safety check

            newTab.classList.add('active');
            newContainer.classList.add('active');

            this.activeSetId = setId;
            const calc = this.calculators.get(setId);
            if (calc instanceof WeaponCalculator) {
                calc.updateSummaryHeader();
            }

            // Define the new handler for the 'Enter' key
            this.activeEnterHandler = (event) => {
                if (event.key === 'Enter' && !event.target.classList.contains('tab-name')) {
                    event.preventDefault();
                    if (calc instanceof WeaponCalculator) {
                        calc.calculateDdoDamage();
                    } else if (calc instanceof SpellCalculator) {
                        calc.calculateSpellDamage();
                    }
                }
            };
            this.activeContainerForEnter = newContainer;
            this.activeContainerForEnter.addEventListener('keydown', this.activeEnterHandler);
        }

        updateComparisonTable() {
            if (!this.comparisonTbody) return;

            // First, find the maximum total average damage among all sets
            let maxDamage = 0;
            if (this.calculators.size > 0) {
                const allDamages = Array.from(this.calculators.values()).map(calc => calc.totalAverageDamage);
                maxDamage = Math.max(...allDamages);
            }
            this.comparisonTbody.innerHTML = ''; // Clear existing rows

            this.calculators.forEach(calc => {
                const row = document.createElement('tr');

                let diffText = 'N/A';
                if (maxDamage > 0) {
                    if (calc.totalAverageDamage === maxDamage) {
                        diffText = `<span class="best-damage-badge">Best</span>`;
                    } else {
                        const diff = ((calc.totalAverageDamage - maxDamage) / maxDamage) * 100;
                        diffText = `${diff.toFixed(1)}%`;
                    }
                }

                // Create cells programmatically to prevent XSS from user-provided tab names.
                const nameCell = document.createElement('td');
                nameCell.textContent = calc.getTabName(); // .textContent is safe
                row.appendChild(nameCell);

                const totalDmgCell = document.createElement('td');
                totalDmgCell.textContent = calc.totalAverageDamage.toFixed(2);
                row.appendChild(totalDmgCell);

                const diffCell = document.createElement('td');
                diffCell.innerHTML = diffText; // Safe because diffText is internally generated ('Best' badge or a number)
                row.appendChild(diffCell);

                if (calc instanceof WeaponCalculator) {
                    const baseCell = document.createElement('td');
                    baseCell.textContent = calc.totalAvgBaseHitDmg.toFixed(2);
                    row.appendChild(baseCell);

                    const sneakCell = document.createElement('td');
                    sneakCell.textContent = calc.totalAvgSneakDmg.toFixed(2);
                    row.appendChild(sneakCell);

                    const imbueCell = document.createElement('td');
                    imbueCell.textContent = calc.totalAvgImbueDmg.toFixed(2);
                    row.appendChild(imbueCell);

                    const unscaledCell = document.createElement('td');
                    unscaledCell.textContent = calc.totalAvgUnscaledDmg.toFixed(2);
                    row.appendChild(unscaledCell);

                    const scaledCell = document.createElement('td');
                    scaledCell.textContent = calc.totalAvgScaledDiceDmg.toFixed(2); // This is totalAvgScaledDiceDmg
                    row.appendChild(scaledCell);

                } else {
                    // For spell calculators, show relevant spell data
                    const avgHitCell = document.createElement('td');
                    avgHitCell.textContent = calc.averageBaseHit.toFixed(2); // Corresponds to Avg Base
                    row.appendChild(avgHitCell);

                    const avgCritCell = document.createElement('td');
                    avgCritCell.textContent = calc.averageCritHit.toFixed(2); // Corresponds to Avg Sneak
                    row.appendChild(avgCritCell);

                    // Add empty cells for the remaining columns (Imbue, Unscaled, Scaled) to keep alignment
                    row.appendChild(document.createElement('td')).textContent = '-'; // Avg Imbue
                    row.appendChild(document.createElement('td')).textContent = '-'; // Avg Unscaled
                    row.appendChild(document.createElement('td')).textContent = '-'; // Avg Scaled
                }

                this.comparisonTbody.appendChild(row);
            });
        }

        addDragAndDropListeners() {
            let draggedTab = null;
            let placeholder = null; // Placeholder element for drop location

            // Use event delegation on the container for drag events
            this.navContainer.addEventListener('dragstart', (e) => {
                const target = e.target.closest('.nav-tab');
                if (target) {
                    draggedTab = target;
                    // Create placeholder
                    placeholder = document.createElement('div');
                    placeholder.className = 'nav-tab-placeholder';
                    placeholder.style.width = `${draggedTab.offsetWidth}px`;
                    placeholder.style.height = `${draggedTab.offsetHeight}px`;

                    setTimeout(() => {
                        draggedTab.classList.add('dragging');
                    }, 0);
                }
            });

            this.navContainer.addEventListener('dragend', (e) => {
                if (draggedTab) {
                    draggedTab.classList.remove('dragging');
                    draggedTab = null;
                    // Clean up placeholder
                    if (placeholder && placeholder.parentNode) {
                        placeholder.parentNode.removeChild(placeholder);
                    }
                    placeholder = null;
                }
            });

            this.navContainer.addEventListener('dragover', (e) => {
                e.preventDefault(); // This is necessary to allow a drop
                if (!placeholder) return;

                const afterElement = this._getDragAfterElement(this.navContainer, e.clientX);
                const firstButton = this.navContainer.querySelector('.nav-action-btn');

                if (afterElement) {
                    // Case 1: Hovering over another tab. Insert placeholder before it.
                    this.navContainer.insertBefore(placeholder, afterElement);
                } else {
                    // Case 2: Not hovering over a tab. This means we are at the end of the tab list.
                    // Insert the placeholder before the first action button.
                    if (firstButton) {
                        this.navContainer.insertBefore(placeholder, firstButton);
                    } else {
                        this.navContainer.appendChild(placeholder); // Fallback
                    }
                }
            });

            this.navContainer.addEventListener('drop', (e) => {
                e.preventDefault();
                if (!draggedTab) return;

                // This calculation needs to be consistent with 'dragover'
                const afterElement = this._getDragAfterElement(this.navContainer, e.clientX);
                const firstButton = this.navContainer.querySelector('.nav-action-btn');
                const draggedSetId = draggedTab.dataset.set;
                const draggedContainer = document.getElementById(`calculator-set-${draggedSetId}`);

                if (afterElement) {
                    // Drop before the 'afterElement'
                    this.navContainer.insertBefore(draggedTab, afterElement);
                    const afterContainer = document.getElementById(`calculator-set-${afterElement.dataset.set}`);
                    this.setsContainer.insertBefore(draggedContainer, afterContainer);
                } else {
                    // Drop at the end (before the first button)
                    if (firstButton) {
                        this.navContainer.insertBefore(draggedTab, firstButton);
                    } else {
                        this.navContainer.appendChild(draggedTab);
                    }
                    this.setsContainer.appendChild(draggedContainer);
                }


                // After reordering, ensure the active tab's class is correctly applied
                // This handles cases where the active tab itself was dragged
                this.switchToSet(this.activeSetId);
                this.saveState(); // Save the new order
            });
        }

        addImportExportListeners() {
            this.exportBtn.addEventListener('click', () => this.showExportModal());
            this.importBtn.addEventListener('click', () => this.showImportModal());
            this.modalCloseBtn.addEventListener('click', () => this.hideModal());
            this.modalBackdrop.addEventListener('click', (e) => {
                if (e.target === this.modalBackdrop) {
                    this.hideModal();
                }
            });

            this.modalCopyBtn.addEventListener('click', () => this.copyToClipboard());
            this.modalLoadBtn.addEventListener('click', () => this.importFromText());

            // Trigger file input when "Load from File" is conceptually clicked
            this.modalFileInput.addEventListener('change', (e) => this.importFromFile(e));

            // Listeners for format toggle
            this.formatJsonBtn.addEventListener('click', () => this.setExportFormat('json'));
            this.formatSummaryBtn.addEventListener('click', () => this.setExportFormat('summary'));
        }

        showExportModal() {
            this.modalTitle.textContent = 'Export Sets';
            this.modalDescription.textContent = 'Copy the text below or save it to a file to import later.';
            this.modalTextarea.value = this.getSetsAsJSON();
            this.modalTextarea.readOnly = true;

            // Show export elements, hide import elements
            this.modalCopyBtn.classList.remove('hidden');
            this.modalSaveFileBtn.classList.remove('hidden');
            this.modalLoadBtn.classList.add('hidden');
            this.modalFileInput.classList.add('hidden');
            document.querySelector('.modal-format-toggle').classList.remove('hidden');

            this.modalSaveFileBtn.onclick = () => this.saveToFile();
            this.modalBackdrop.classList.remove('hidden');
            this.setExportFormat('json'); // Default to JSON
        }

        setExportFormat(format) {
            this.formatJsonBtn.classList.toggle('active', format === 'json');
            this.formatSummaryBtn.classList.toggle('active', format === 'summary');
            this.modalTextarea.value = format === 'json' ? this.getSetsAsJSON() : this.getSetsAsSummary();
            this.modalDescription.textContent = format === 'json' ? 'Copy the text below or save it to a file to import later.' : 'A human-readable summary for sharing or saving as a .txt file.';
        }

        showImportModal() {
            this.modalTitle.textContent = 'Import Sets';
            this.modalDescription.textContent = 'Paste set data into the text area and click "Load", or upload a file.';
            this.modalTextarea.value = '';
            this.modalTextarea.readOnly = false;
            this.modalTextarea.placeholder = 'Paste your exported set data here...';

            // Show import elements, hide export elements
            this.modalCopyBtn.classList.add('hidden');
            this.modalSaveFileBtn.classList.add('hidden');
            this.modalLoadBtn.classList.remove('hidden');
            // We don't show the file input directly, but we'll trigger it.
            document.querySelector('.modal-format-toggle').classList.add('hidden');
            // Let's repurpose the "Save to File" button to be "Load from File"
            this.modalSaveFileBtn.textContent = 'Load from File';
            this.modalSaveFileBtn.classList.remove('hidden');
            this.modalSaveFileBtn.onclick = () => this.modalFileInput.click(); // Re-route click

            this.modalBackdrop.classList.remove('hidden');
        }

        hideModal() {
            this.modalBackdrop.classList.add('hidden');
            // Reset the repurposed button
            this.modalSaveFileBtn.textContent = 'Save to File';
            this.modalSaveFileBtn.onclick = () => this.saveToFile();
        }

        getSetsAsJSON() {
            const stateToSave = [];
            const orderedTabs = this.navContainer.querySelectorAll('.nav-tab');
            orderedTabs.forEach(tab => {
                const setId = parseInt(tab.dataset.set, 10);
                const calc = this.calculators.get(setId);
                if (calc) {
                    const state = calc.getState();
                    state.tabName = calc.getTabName();
                    state.setId = calc.setId;
                    state.type = calc instanceof SpellCalculator ? 'spell' : 'weapon';
                    stateToSave.push(state);
                }
            });
            return JSON.stringify(stateToSave, null, 2); // Pretty-print the JSON
        }

        getSetsAsSummary() {
            let summary = `DDO Damage Calculator Export\nGenerated: ${new Date().toLocaleString()}\n\n`;
            const orderedTabs = this.navContainer.querySelectorAll('.nav-tab');

            orderedTabs.forEach(tab => {
                const setId = parseInt(tab.dataset.set, 10);
                const calc = this.calculators.get(setId);
                if (!calc) return;

                const state = calc.getState();
                summary += `========================================\n`;
                summary += `  Set: ${calc.getTabName()}\n`;
                summary += `========================================\n\n`;

                if (calc instanceof WeaponCalculator) {
                    summary += `--- Base Damage ---\n`;
                    summary += `Weapon Dice [W]: ${state['weapon-dice'] || 0}\n`;
                    summary += `Damage: ${state['weapon-damage'] || '0'} + ${state['bonus-base-damage'] || 0}\n\n`;

                    summary += `--- Critical Profile ---\n`;
                    summary += `Threat Range: ${state['crit-threat'] || '20'}\n`;
                    summary += `Multiplier: x${state['crit-multiplier'] || 2}\n`;
                    summary += `Seeker: +${state['seeker-damage'] || 0}\n`;
                    summary += `19-20 Multiplier: +${state['crit-multiplier-19-20'] || 0}\n\n`;

                    summary += `--- Hit/Miss Profile ---\n`;
                    summary += `Miss on Roll <=: ${state['miss-threshold'] || 1}\n`;
                    summary += `Graze on Roll <=: ${state['graze-threshold'] || 0}\n`;
                    summary += `Graze Damage: ${state['graze-percent'] || 0}%\n\n`;

                    summary += `--- Unscaled Damage ---\n`;
                    let i = 1;
                    while (state.hasOwnProperty(`unscaled-damage-${i}`)) {
                        const damage = state[`unscaled-damage-${i}`] || '0';
                        if (damage && damage !== '0') {
                            const proc = state[`unscaled-proc-chance-${i}`] || 100;
                            const multi = state[`unscaled-doublestrike-${i}`] ? 'Yes' : 'No';
                            const onCrit = state[`unscaled-on-crit-${i}`] ? ', On Crit Only' : '';
                            summary += `Source ${i}: ${damage} @ ${proc}% Proc, Multi-Strike: ${multi}${onCrit}\n`;
                        }
                        i++;
                    }
                    summary += `Melee/Ranged Power: ${state['melee-power'] || 0}\n`;
                    summary += `Spell Power: ${state['spell-power'] || 0}\n`;
                    summary += `Multi-Strike: ${state['doublestrike'] || 0}% (${state['is-doubleshot'] ? 'Doubleshot' : 'Doublestrike'})\n\n`;
                    summary += `Archer's Focus: ${state['archers-focus'] || 0} stacks (+${(state['archers-focus'] || 0) * 5} RP)\n\n`;

                    summary += `--- Sneak Attack ---\n`;
                    summary += `Damage: ${state['sneak-attack-dice'] || 0}d6 + ${state['sneak-bonus'] || 0}\n\n`;

                    summary += `--- Imbue Dice ---\n`;
                    summary += `Dice: ${state['imbue-dice-count'] || 0}d${state['imbue-die-type'] || 6}\n`;
                    summary += `Scaling: ${state['imbue-scaling'] || 100}% of ${state['imbue-uses-spellpower'] ? 'Spell Power' : 'Melee Power'}\n\n`;

                    summary += `--- AVERAGES ---\n`;
                    summary += `Total Avg Damage: ${calc.totalAverageDamage.toFixed(2)}\n`;
                    summary += `Avg Base: ${calc.totalAvgBaseHitDmg.toFixed(2)}, Avg Sneak: ${calc.totalAvgSneakDmg.toFixed(2)}, Avg Imbue: ${calc.totalAvgImbueDmg.toFixed(2)}, Avg Unscaled: ${calc.totalAvgUnscaledDmg.toFixed(2)}\n\n\n`;
                } else if (calc instanceof SpellCalculator) {
                    // Reconstruct structured data from flat state for summary generation
                    const spellPowerProfiles = [];
                    const spellDamageSources = [];
                    for (const key in state) {
                        if (key.startsWith('spell-power-type-')) {
                            const id = key.substring('spell-power-type-'.length);
                            spellPowerProfiles.push({
                                id: parseInt(id, 10),
                                type: state[key] || 'Unnamed',
                                spellPower: state[`spell-power-${id}`] || 0,
                                critChance: state[`spell-crit-chance-${id}`] || 0,
                                critDamage: state[`spell-crit-damage-${id}`] || 0,
                            });
                        } else if (key.startsWith('spell-name-')) {
                            const id = key.substring('spell-name-'.length);
                            const source = {
                                id: parseInt(id, 10),
                                name: state[key] || `Source ${id}`,
                                base: state[`spell-damage-${id}`] || '0',
                                clScaled: state[`spell-cl-scaling-${id}`] || '0',
                                casterLevel: state[`caster-level-${id}`] || 0,
                                hitCount: state[`spell-hit-count-${id}`] || 1,
                                additionalScalings: [],
                            };
                            // Now, find additional scalings for this source
                            for (const subKey in state) {
                                if (subKey.startsWith(`additional-scaling-base-${id}-`)) {
                                    const scalingId = subKey.substring(`additional-scaling-base-${id}-`.length);
                                    source.additionalScalings.push({
                                        base: state[subKey] || '0',
                                        clScaled: state[`additional-scaling-cl-${id}-${scalingId}`] || '0',
                                        profileId: state[`additional-scaling-sp-select-${id}-${scalingId}`] || 1,
                                    });
                                }
                            }
                            spellDamageSources.push(source);
                        }
                    }
                    // Sort by ID to ensure correct order
                    spellPowerProfiles.sort((a, b) => a.id - b.id);
                    spellDamageSources.sort((a, b) => a.id - b.id);

                    summary += `--- Metamagics & Boosts ---\n`;
                    const activeToggles = [];
                    if (state['metamagic-empower']) activeToggles.push('Empower');
                    if (state['metamagic-maximize']) activeToggles.push('Maximize');
                    if (state['metamagic-intensify']) activeToggles.push('Intensify');
                    if (state['boost-wellspring']) activeToggles.push('Wellspring of Power');
                    if (state['boost-night-horrors']) activeToggles.push('Night Horrors');
                    summary += activeToggles.length > 0 ? activeToggles.join(', ') + '\n\n' : 'None\n\n';

                    summary += `--- Spell Power Profiles ---\n`;
                    spellPowerProfiles.forEach(p => {
                        summary += `Profile ${p.id} (${p.type}): ${p.spellPower} SP, ${p.critChance}% Chance, +${p.critDamage}% Damage\n`;
                    });
                    summary += `\n`;

                    summary += `--- Spell Damage Sources ---\n`;
                    spellDamageSources.forEach(s => {
                        summary += `Source "${s.name}" (x${s.hitCount} hits):\n`;
                        summary += `  - Base: ${s.base} + (${s.clScaled} per CL) @ CL ${s.casterLevel}\n`;
                        s.additionalScalings.forEach(as => {
                            summary += `  - Extra: ${as.base} + (${as.clScaled} per CL) [Uses SP ${as.profileId}]\n`;
                        });
                    });
                    summary += `\n`;

                    summary += `--- AVERAGES ---\n`;
                    summary += `Total Average Damage: ${calc.totalAverageDamage.toFixed(2)}\n\n\n`;
                }
            });

            return summary;
        }

        copyToClipboard() {
            navigator.clipboard.writeText(this.modalTextarea.value).then(() => {
                alert('Set data copied to clipboard!');
            }).catch(err => {
                console.error('Failed to copy text: ', err);
                alert('Failed to copy. Please copy manually from the text box.');
            });
        }

        saveToFile() {
            const isJson = this.formatJsonBtn.classList.contains('active');
            const data = this.modalTextarea.value;
            const fileType = isJson ? 'application/json' : 'text/plain';
            const fileName = isJson ? 'ddo-calc-sets.json' : 'ddo-calc-summary.txt';
            const blob = new Blob([data], { type: fileType });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        importFromText() {
            const jsonString = this.modalTextarea.value;
            if (!jsonString.trim()) {
                alert('Text area is empty. Please paste your set data.');
                return;
            }
            this.loadSetsFromJSON(jsonString);
        }

        importFromFile(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (e) => {
                const jsonString = e.target.result;
                this.loadSetsFromJSON(jsonString);
            };
            reader.onerror = () => {
                alert('Error reading file.');
            };
            reader.readAsText(file);

            // Reset file input to allow loading the same file again
            event.target.value = '';
        }

        // This method is called when a user directly changes an input,
        // or when an undo/redo action is completed.
        // It should not be called during the loading process.
        saveState() {
            if (this.isLoading) return; // Don't save while loading


            const stateToSave = [];
            // Get tabs in their current DOM order to save the correct sequence
            const orderedTabs = this.navContainer.querySelectorAll('.nav-tab');
            orderedTabs.forEach(tab => {
                const setId = parseInt(tab.dataset.set, 10);
                const calc = this.calculators.get(setId);
                if (!calc) return; // Skip if calculator doesn't exist for some reason

                const state = calc.getState();
                state.tabName = calc.getTabName(); // Save the tab name
                state.setId = calc.setId; // Save the ID
                state.type = calc instanceof SpellCalculator ? 'spell' : 'weapon';
                stateToSave.push(state);
            })
            sessionStorage.setItem('calculatorState', JSON.stringify(stateToSave));
            sessionStorage.setItem('activeSetId', this.activeSetId); // Save the active set ID
        }


        loadState() {
            const jsonString = sessionStorage.getItem('calculatorState');
            const activeSetId = sessionStorage.getItem('activeSetId'); // Get the saved active ID
            if (this.loadSetsFromJSON(jsonString, activeSetId)) {
                return true;
            }
            return false;
        }

        loadSetsFromJSON(jsonString, activeSetIdToLoad) {
            if (!jsonString) return false;

            let savedStates;
            try {
                savedStates = JSON.parse(jsonString);
                if (!Array.isArray(savedStates) || savedStates.length === 0) {
                    throw new Error("Data is not a valid array of sets.");
                }
            } catch (error) {
                console.error("Failed to parse set data:", error);
                alert("Import failed. The provided data is not valid JSON or is incorrectly formatted.");
                return false;
            }

            this.isLoading = true;

            // Clear existing sets
            this.calculators.clear();
            this.setsContainer.innerHTML = '';
            this.navContainer.querySelectorAll('.nav-tab').forEach(tab => tab.remove()); // Only remove tabs, not other buttons
            this.nextSetId = 1; // Reset counter

            savedStates.forEach((state) => {
                this.addNewSet(state.type, state.setId);
                const newCalc = this.calculators.get(state.setId);
                newCalc?.setState(state);
                newCalc?.setTabName(state.tabName);
                if (newCalc instanceof WeaponCalculator) {
                    newCalc.updateSummaryHeader();
                    newCalc.calculateDdoDamage(); // Recalculate after setting state
                } else if (newCalc instanceof SpellCalculator) {
                    newCalc.calculateSpellDamage(); // Recalculate for spells too
                }
            });

            // Switch to the saved active set, or default to the first one if not found
            const targetSetId = activeSetIdToLoad ? parseInt(activeSetIdToLoad, 10) : savedStates[0]?.setId;
            this.switchToSet(targetSetId || 1);
            this.isLoading = false;
            this.updateComparisonTable(); // Populate comparison table after loading
            this.hideModal();
            this.saveState(); // Save the newly loaded state to sessionStorage
            return true;
        }

        addUndoRedoListeners() {
            document.addEventListener('keydown', (e) => {
                // Check for Ctrl+Z (Undo)
                if (e.ctrlKey && e.key.toLowerCase() === 'z' && !e.shiftKey) {
                    e.preventDefault();
                    this.undo();
                }
                // Check for Ctrl+Y or Ctrl+Shift+Z (Redo)
                if ((e.ctrlKey && e.key.toLowerCase() === 'y') || (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'z')) {
                    e.preventDefault();
                    this.redo();
                }
            });
        }

        recordAction(action) {
            this.undoStack.push(action);
            // A new action clears the redo stack
            this.redoStack = [];
        }

        undo() {
            if (this.undoStack.length === 0) return;

            const action = this.undoStack.pop();
            this.isLoading = true; // Prevent saving state during undo

            if (action.type === 'REMOVE_SET') {
                // To undo a remove, we add the set back
                this.recreateSet(action.setId, action.state, action.index);
                // The redo action is to remove it again
                this.redoStack.push({ ...action, type: 'REMOVE_SET' });
            } else if (action.type === 'ADD_SET') {
                // To undo an add, we remove the set
                this.removeSet(action.setId, true); // Use isLoading to prevent recording
                // The redo action is to add it back
                this.redoStack.push({ ...action, type: 'ADD_SET' }); // The 'add' action doesn't have a state to copy
            } else if (action.type === 'VALUE_CHANGE' || action.type === 'RENAME') {
                // To undo a value change, we apply the old value.
                const calc = this.calculators.get(action.setId);
                if (action.type === 'RENAME') {
                    calc?.setTabName(action.oldValue);
                } else {
                    calc?.setState({ [action.key]: action.oldValue });
                }
                this.redoStack.push(action);
            } else if (action.type === 'ADD_DYNAMIC_ROW') {
                const calc = this.calculators.get(action.setId);
                calc?.removeDynamicRow(action.rowId, action.rowType);
                this.redoStack.push(action);
            } else if (action.type === 'REMOVE_DYNAMIC_ROW') {
                const calc = this.calculators.get(action.setId);
                calc?.recreateDynamicRow(action.rowId, action.rowType, action.rowData, action.rowIndex);
                this.redoStack.push(action);
            }

            this.isLoading = false;
            this.saveState();
            this.updateComparisonTable();
        }

        redo() {
            if (this.redoStack.length === 0) return;

            const action = this.redoStack.pop();
            this.isLoading = true; // Prevent saving state during redo

            if (action.type === 'REMOVE_SET') {
                this.removeSet(action.setId, true);
                this.undoStack.push({ ...action, type: 'REMOVE_SET' });
            } else if (action.type === 'ADD_SET') {
                this.recreateSet(action.setId, action.state, action.index);
                this.undoStack.push({ ...action, type: 'ADD_SET' });
            } else if (action.type === 'VALUE_CHANGE' || action.type === 'RENAME') {
                const calc = this.calculators.get(action.setId);
                if (action.type === 'RENAME') {
                    calc?.setTabName(action.newValue);
                } else {
                    calc?.setState({ [action.key]: action.newValue });
                }
                this.undoStack.push(action);
            } else if (action.type === 'ADD_DYNAMIC_ROW') {
                const calc = this.calculators.get(action.setId);
                calc?.recreateDynamicRow(action.rowId, action.rowType, action.rowData, action.rowIndex);
                this.undoStack.push(action);
            } else if (action.type === 'REMOVE_DYNAMIC_ROW') {
                const calc = this.calculators.get(action.setId);
                calc?.removeDynamicRow(action.rowId, action.rowType);
                this.undoStack.push(action);
            }

            this.isLoading = false;
            this.saveState();
            this.updateComparisonTable();
        }

        /**
         * Helper function to determine the element a dragged tab should be placed before.
         * @param {HTMLElement} container - The parent container of the draggable elements.
         * @param {number} x - The current X coordinate of the mouse during drag.
         * @returns {HTMLElement|null} The element to drop before, or null if dropping at the end.
         */
        _getDragAfterElement(container, x) {
            const draggableElements = [...container.querySelectorAll('.nav-tab:not(.dragging)')];
            return draggableElements.reduce((closest, child) => {
                const box = child.getBoundingClientRect();
                const offset = x - box.left - box.width / 2;
                return (offset < 0 && offset > closest.offset) ? { offset: offset, element: child } : closest;
            }, { offset: Number.NEGATIVE_INFINITY }).element;
        }

        getTemplateHTML() {
            const templateNode = document.getElementById('calculator-set-template');
            return templateNode ? templateNode.innerHTML : '';
        }

        findNextAvailableId() {
            let id = 1;
            while (this.calculators.has(id)) {
                id++;
            }
            if (id >= this.nextSetId) this.nextSetId = id + 1; // Update nextSetId here
            return id;
        }
    }

    // --- Theme Toggler Logic ---
    const themeToggleCheckbox = document.getElementById('theme-toggle-checkbox');
    const body = document.body;
    const currentTheme = localStorage.getItem('theme');

    // Apply saved theme on load
    if (themeToggleCheckbox) {
        if (currentTheme === 'dark') {
            body.classList.add('dark-mode');
            themeToggleCheckbox.checked = true;
        }

        themeToggleCheckbox.addEventListener('change', () => {
            body.classList.toggle('dark-mode');
            let theme = 'light';
            if (body.classList.contains('dark-mode')) {
                theme = 'dark';
            }
            localStorage.setItem('theme', theme);
        });
    }

    // --- Instantiate Manager ---
    const manager = new CalculatorManager();
});