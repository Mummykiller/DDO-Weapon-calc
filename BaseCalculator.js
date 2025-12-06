import { serializeForm, deserializeForm } from './utils.js';

export class BaseCalculator {
    constructor(setId, manager, name) {
        this.setId = setId;
        this.manager = manager;
        this.name = name;
        this.idSuffix = `-set${setId}`;
        this.container = document.getElementById(`calculator-set-${this.setId}`);
        this.totalAverageDamage = 0;

        // Create a hidden span for measuring text width
        this._measurementSpan = document.createElement('span');
        this._measurementSpan.style.position = 'absolute';
        this._measurementSpan.style.visibility = 'hidden';
        this._measurementSpan.style.whiteSpace = 'nowrap';
        document.body.appendChild(this._measurementSpan);
    }

    getElements() {
        // This method will be implemented by the child classes
    }

    addEventListeners() {
        // This method is intended to be overridden by child classes
    }

    removeEventListeners() {
        // Event listeners are on the container, which gets removed, 
        // so no specific removal is needed here with the current event delegation model.
    }

    _resizeInput(inputElement) {
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

        let desiredWidth = this._measurementSpan.offsetWidth + paddingLeft + paddingRight + borderWidthLeft + borderWidthRight + 4;

        const minWidth = parseFloat(computedStyle.minWidth) || 50;

        inputElement.style.width = `${Math.max(minWidth, desiredWidth)}px`;
    }

    _initializeAdaptiveInputs() {
        if (!this.container) return;

        this.container.addEventListener('input', (e) => {
            if (e.target.classList.contains('adaptive-text-input')) {
                this._resizeInput(e.target);
            }
        });

        this.resizeAllAdaptiveInputs();
    }

    resizeAllAdaptiveInputs() {
        this.container?.querySelectorAll('.adaptive-text-input').forEach(input => this._resizeInput(input));
    }

    getState() {
        return this.state;
    }

    setState(state) {
        deserializeForm(this.container, this.idSuffix, state);
        this.resizeAllAdaptiveInputs();
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

    handleInputChange() {
        // This method will be implemented by child classes
    }
}
