/**
 * Parses a dice notation string (e.g., "2d6+3") and returns its average value.
 * Also handles flat numbers and subtraction.
 * @param {string} diceString - The string to parse, like "1d8", "3d6-2", or "4.5".
 * @returns {number} The calculated average damage.
 */
export function parseDiceNotation(diceString) {
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

            if (isNaN(numDice)) numDice = 1; // Default for invalid strings like "ad6"

            const numSides = parseFloat(parts[1]); // Use parseFloat to handle dice like 'd100' or 'd4.5' if ever needed

            if (isNaN(numSides) || numSides <= 0) continue; // Invalid sides, skip

            // Average of one die is (sides + 1) / 2. Multiply by the number of a dice.
            totalAverage += numDice * (numSides + 1) / 2;
        } else {
            // If not a die roll, it's a flat number (e.g., "5" or "-2")
            totalAverage += (parseFloat(term) || 0);
        }
    }
    return totalAverage;
}
