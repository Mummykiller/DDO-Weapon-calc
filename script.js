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

            this.activeSetId = 1;
            this.nextSetId = 1; // Moved nextSetId into CalculatorManager

            this.addSetBtn.addEventListener('click', () => this.addNewSet());
            this.addSpellSetBtn.addEventListener('click', () => this.addNewSpellSet());
            // Try to load state. If it fails (e.g., first visit), create the initial set.
            if (!this.loadState()) {
                this.addNewSet(1);
                this.calculators.get(1)?.calculateDdoDamage(); // Perform initial calculation
            }

            this.addDragAndDropListeners();
            this.addImportExportListeners();
            this.addUndoRedoListeners();
        }


        addNewSet(setIdToUse = null, index = -1) {
            if (this.calculators.size >= 6) {
                alert("You have reached the maximum of 6 weapon sets.");
                return;
            }

            // If this is a user action (not part of loading or undo), record it.
            if (setIdToUse === null) {
                // We need to know the ID before creating it.
                const futureSetId = this.findNextAvailableId();
                this.recordAction({ type: 'add', setId: futureSetId });
            }

            // Get the state of the currently active set to copy it
            const activeCalc = this.calculators.get(this.activeSetId);
            const stateToCopy = (setIdToUse === null && activeCalc instanceof WeaponCalculator) ? activeCalc.getState() : null; // Only copy state if creating a new set from an existing one

            let newSetId;
            if (setIdToUse !== null) {
                newSetId = setIdToUse;
                // Ensure the global nextSetId counter is always ahead of the largest known ID.
                // This prevents ID collisions when adding a new set after loading a specific set.
                if (newSetId >= this.nextSetId) {
                    this.nextSetId = newSetId + 1;
                }
            } else {
                newSetId = this.findNextAvailableId();
            }

            // Get the inner HTML of the template set (calculator-set-1)
            const templateNode = document.getElementById('calculator-set-template').content.cloneNode(true);

            // Replace IDs and 'for' attributes within the inner HTML
            let modifiedInnerHtml = templateNode.firstElementChild.outerHTML.replace(/\s(id)=\"([^\"]+)\"/g, (match, attr, id) => {
                return ` id="${id}-set${newSetId}"`;
            });
            modifiedInnerHtml = modifiedInnerHtml.replace(/for=\"([^\"]+)\"/g, (match, id) => {
                return `for="${id}-set${newSetId}"`;
            });

            // Create the new container div and set its properties
            const newSetContainer = document.createElement('div');
            newSetContainer.id = `calculator-set-${newSetId}`; // e.g., calculator-set-2
            newSetContainer.className = 'calculator-container calculator-set'; // Copy classes from template
            newSetContainer.innerHTML = modifiedInnerHtml; // Set the modified inner HTML

            // Append the new container to the DOM at the correct index
            const allContainers = this.setsContainer.querySelectorAll('.calculator-set');
            if (index !== -1 && index < allContainers.length) {
                this.setsContainer.insertBefore(newSetContainer, allContainers[index]);
            } else {
                this.setsContainer.appendChild(newSetContainer);
            }

            // Create and add the new tab at the correct index
            const tab = this.createTab(newSetId);
            const allTabs = this.navContainer.querySelectorAll('.nav-tab');
            if (index !== -1 && index < allTabs.length) {
                this.navContainer.insertBefore(tab, allTabs[index]);
            } else {
                const firstButton = this.navContainer.querySelector('.nav-action-btn');
                this.navContainer.insertBefore(tab, firstButton || null);
            }

            this.calculators.set(newSetId, new WeaponCalculator(newSetId, this, `Set ${newSetId}`));
            const newCalc = this.calculators.get(newSetId);

            // If we have a state to copy, apply it to the new set
            if (stateToCopy && newCalc) {
                newCalc.setState(stateToCopy);
            }

            this.switchToSet(newSetId);
            if (!this.isLoading) {
                this.saveState();
            }
            this.updateComparisonTable();
        }

        addNewSpellSet(setIdToUse = null, index = -1) {
            if (this.calculators.size >= 6) {
                alert("You have reached the maximum of 6 sets.");
                return;
            }

            const activeCalc = this.calculators.get(this.activeSetId);
            const stateToCopy = (setIdToUse === null && activeCalc instanceof SpellCalculator) ? activeCalc.getState() : null;

            let newSetId;
            if (setIdToUse !== null) {
                newSetId = setIdToUse;
                if (newSetId >= this.nextSetId) {
                    this.nextSetId = newSetId + 1;
                }
            } else {
                newSetId = this.findNextAvailableId();
            }

            const templateNode = document.getElementById('spell-calculator-template').content.cloneNode(true);

            let modifiedInnerHtml = templateNode.firstElementChild.outerHTML.replace(/\s(id)=\"([^\"]+)\"/g, (match, attr, id) => {
                return ` id="${id}-set${newSetId}"`;
            });
            modifiedInnerHtml = modifiedInnerHtml.replace(/for=\"([^\"]+)\"/g, (match, id) => {
                return `for="${id}-set${newSetId}"`;
            });

            const newSetContainer = document.createElement('div');
            newSetContainer.id = `calculator-set-${newSetId}`;
            newSetContainer.className = 'calculator-container calculator-set';
            newSetContainer.innerHTML = modifiedInnerHtml;

            // Append the new container to the DOM at the correct index
            const allContainers = this.setsContainer.querySelectorAll('.calculator-set');
            if (index !== -1 && index < allContainers.length) {
                this.setsContainer.insertBefore(newSetContainer, allContainers[index]);
            } else {
                this.setsContainer.appendChild(newSetContainer);
            }

            // Create and add the new tab at the correct index
            const tab = this.createTab(newSetId);
            tab.classList.add('spell-tab-indicator');
            const allTabs = this.navContainer.querySelectorAll('.nav-tab');
            if (index !== -1 && index < allTabs.length) {
                this.navContainer.insertBefore(tab, allTabs[index]);
            } else {
                const firstButton = this.navContainer.querySelector('.nav-action-btn');
                this.navContainer.insertBefore(tab, firstButton || null);
            }


            this.calculators.set(newSetId, new SpellCalculator(newSetId, this, `Spell Set ${newSetId}`));
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
            tabNameSpan.addEventListener('blur', () => {
                const calc = this.calculators.get(setId);
                if (calc) {
                    if (calc instanceof WeaponCalculator) {
                        calc.handleInputChange();
                    } else {
                        calc.calculateSpellDamage();
                    }
                    if (this.activeSetId === setId) {
                        calc.updateSummaryHeader();
                    }
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
                            calc.calculateSpellDamage();
                        }
                        if (this.activeSetId === setId) calc.updateSummaryHeader();
                    }
                }
            });
        }

        recreateSet(setId, state, index) {
            if (state.type === 'spell') {
                this.addNewSpellSet(setId, index);
            } else {
                this.addNewSet(setId, index);
            }
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
            // Deactivate all
            document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.calculator-set').forEach(s => s.classList.remove('active'));

            // Activate the selected one
            document.querySelector(`.nav-tab[data-set="${setId}"]`).classList.add('active');
            document.getElementById(`calculator-set-${setId}`).classList.add('active');
            this.activeSetId = setId;
            const calc = this.calculators.get(setId);
            if (calc instanceof WeaponCalculator) {
                calc.updateSummaryHeader();
            }
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
            this.modalSaveFileBtn.addEventListener('click', () => this.saveToFile());
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

                    summary += `--- Sneak Attack ---\n`;
                    summary += `Damage: ${state['sneak-attack-dice'] || 0}d6 + ${state['sneak-bonus'] || 0}\n\n`;

                    summary += `--- Imbue Dice ---\n`;
                    summary += `Dice: ${state['imbue-dice-count'] || 0}d${state['imbue-die-type'] || 6}\n`;
                    summary += `Scaling: ${state['imbue-scaling'] || 100}% of ${state['imbue-uses-spellpower'] ? 'Spell Power' : 'Melee Power'}\n\n`;

                    summary += `--- AVERAGES ---\n`;
                    summary += `Total Avg Damage: ${calc.totalAverageDamage.toFixed(2)}\n`;
                    summary += `Avg Base: ${calc.totalAvgBaseHitDmg.toFixed(2)}, Avg Sneak: ${calc.totalAvgSneakDmg.toFixed(2)}, Avg Imbue: ${calc.totalAvgImbueDmg.toFixed(2)}, Avg Unscaled: ${calc.totalAvgUnscaledDmg.toFixed(2)}\n\n\n`;
                } else if (calc instanceof SpellCalculator) {
                    summary += `--- Spell Properties ---\n`;
                    if (state.spellDamageSources && state.spellDamageSources.length > 0) {
                        state.spellDamageSources.forEach((source, index) => {
                            summary += `  Spell ${index + 1} (${source.name || 'Unnamed'}):\n`;
                            summary += `    Base Damage: ${source.base || '0'}\n`;
                            summary += `    CL Scaled: ${source.clScaled || '0'}\n`;
                            // Also include calculated avg hit and avg crit for this spell if available from the calculator instance
                            if (calc.individualSpellDamages && calc.individualSpellDamages[index]) {
                                summary += `    Avg Hit (pre-crit): ${calc.individualSpellDamages[index].averageHit.toFixed(2)}\n`;
                                summary += `    Avg Crit: ${calc.individualSpellDamages[index].averageCrit.toFixed(2)}\n`;
                            }
                        });
                    }
                    summary += `Caster Level: ${state['caster-level'] || 0}\n`;
                    summary += `Spell Power: ${state['spell-power'] || 0}\n`;
                    summary += `Crit Chance: ${state['spell-crit-chance'] || 0}%\n`;
                    summary += `Crit Damage: ${state['spell-crit-damage'] || 0}%\n`;
        

                    summary += `--- AVERAGES ---\n`;
                    summary += `Total Avg Damage: ${calc.totalAverageDamage.toFixed(2)}\n\n\n`;
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
                if (state.type === 'spell') {
                    this.addNewSpellSet(state.setId);
                } else {
                    this.addNewSet(state.setId);
                }
                const newCalc = this.calculators.get(state.setId);
                newCalc?.setState(state);
                newCalc?.setTabName(state.tabName);
                if (newCalc instanceof WeaponCalculator) {
                    newCalc.updateSummaryHeader();
                }
            });

            // Switch to the saved active set, or default to the first one if not found
            const targetSetId = activeSetIdToLoad ? parseInt(activeSetIdToLoad, 10) : savedStates[0]?.setId;
            this.switchToSet(targetSetId || 1);
            this.isLoading = false;
            this.updateComparisonTable(); // Populate comparison table after loading
            this.hideModal();
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
                const calc = this.calculators.get(action.setId);
                calc?.applyValueChange(action.inputId || `tab-name-${action.setId}`, action.oldValue);
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
                calc?.applyValueChange(action.inputId || `tab-name-${action.setId}`, action.newValue);
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
        // Global Enter key listener
        addGlobalEnterListener() {
            document.addEventListener('keyup', (event) => {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    const activeCalc = this.calculators.get(this.activeSetId);
                    if (activeCalc instanceof WeaponCalculator) {
                        activeCalc.calculateDdoDamage();
                    } else if (activeCalc instanceof SpellCalculator) {
                        activeCalc.calculateSpellDamage();
                    }
                }
            });
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
    manager.addGlobalEnterListener(); // Add the global listener after manager is ready
});