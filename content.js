/**
 * NoHands Content Script
 * Injected into web pages to fill form fields automatically
 */

// Guard against double injection (manifest + chrome.scripting)
if (window.__nohandsInjected) {
  // Already injected, skip
} else {
  window.__nohandsInjected = true;

// Store the last right-clicked element for context menu
let lastContextMenuTarget = null;

// Store last fill request for MutationObserver re-fill
let lastFillData = null;
let lastFillMapping = null;
let lastFillCustomFields = null;
let fillObserver = null;

// Track right-click on any element
document.addEventListener('contextmenu', (event) => {
  lastContextMenuTarget = event.target;
}, true);

// Listen for messages from the popup and background script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'fillForm') {
    console.log('NoHands: Received fillForm message', request);

    // Store for MutationObserver re-fill
    lastFillData = request.data;
    lastFillMapping = request.mapping;
    lastFillCustomFields = request.customFields || null;

    const result = performFill(request.data, request.mapping, request.customFields);

    // Start observing for dynamically loaded content (ASP.NET UpdatePanels, etc.)
    startFillObserver();

    sendResponse(result);
  } else if (request.action === 'copyInputName') {
    if (lastContextMenuTarget) {
      const inputName = lastContextMenuTarget.getAttribute('name');

      if (inputName) {
        // Copy to clipboard
        navigator.clipboard.writeText(inputName).then(() => {
          console.log('NoHands: Input name copied:', inputName);

          // Show visual feedback
          showCopyNotification(lastContextMenuTarget, inputName);

          sendResponse({ success: true, inputName: inputName });
        }).catch(err => {
          console.error('NoHands: Failed to copy:', err);
          sendResponse({ success: false, error: err.message });
        });
      } else {
        console.warn('NoHands: Element has no name attribute');
        sendResponse({ success: false, error: 'No name attribute' });
      }
    } else {
      console.warn('NoHands: No target element');
      sendResponse({ success: false, error: 'No target element' });
    }
    return true; // Keep the message channel open for async response
  }
  return true; // Keep the message channel open for async response
});

/**
 * Perform fill with data, mapping and optional custom fields
 */
function performFill(data, mapping, customFields) {
  const result = fillFormFields(data, mapping);
  if (customFields && typeof customFields === 'object') {
    const customResult = fillCustomFields(customFields);
    result.filledCount += customResult.filledCount;
    if (customResult.filled.length) result.filled.push(...customResult.filled);
    if (customResult.errors && customResult.errors.length) {
      result.errors = (result.errors || []).concat(customResult.errors);
      result.error = result.errors.slice(0, 3).join(', ');
    }
    result.success = result.filledCount > 0;
  }
  return result;
}

/**
 * MutationObserver for dynamically loaded ASP.NET content.
 * Watches for new form elements and re-fills them.
 */
function startFillObserver() {
  // Disconnect any previous observer
  if (fillObserver) {
    fillObserver.disconnect();
    fillObserver = null;
  }

  if (!lastFillData || !lastFillMapping) return;

  let retryCount = 0;
  const maxRetries = 10;
  let debounceTimer = null;

  fillObserver = new MutationObserver((mutations) => {
    // Check if any new form elements were added
    let hasNewInputs = false;
    for (const mutation of mutations) {
      for (const node of mutation.addedNodes) {
        if (node.nodeType === Node.ELEMENT_NODE) {
          if ((node.matches && node.matches('input, select, textarea, form')) ||
              (node.querySelector && node.querySelector('input, select, textarea'))) {
            hasNewInputs = true;
            break;
          }
        }
      }
      if (hasNewInputs) break;
    }

    if (hasNewInputs && retryCount < maxRetries) {
      retryCount++;
      // Debounce — wait for ASP.NET to finish rendering batch of elements
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        console.log(`NoHands: MutationObserver detected new inputs (retry ${retryCount}/${maxRetries}), re-filling...`);
        performFill(lastFillData, lastFillMapping, lastFillCustomFields);
      }, 300);
    }
  });

  fillObserver.observe(document.body || document.documentElement, {
    childList: true,
    subtree: true
  });

  // Auto-disconnect after 30 seconds
  setTimeout(() => {
    if (fillObserver) {
      fillObserver.disconnect();
      fillObserver = null;
      console.log('NoHands: MutationObserver auto-disconnected after 30s');
    }
  }, 30000);
}

/**
 * Fill form fields based on data and mapping
 * @param {Object} data - Parsed Excel data
 * @param {Object} mapping - Column to input name mapping
 * @returns {Object} - Result with success status and filled count
 */
function fillFormFields(data, mapping) {
  let filledCount = 0;
  const errors = [];
  const filled = [];

  console.log('NoHands: Starting to fill form fields', { data, mapping });

  // For each mapping defined
  for (const [columnName, inputNames] of Object.entries(mapping)) {
    // Get value from Excel data
    const value = data[columnName];
    if (value === undefined || value === null) {
      console.log(`NoHands: Skipping ${columnName} - no value in data`);
      continue;
    }

    // Convert inputNames to array if it's a string
    let inputNamesArray = inputNames;
    if (typeof inputNames === 'string') {
      inputNamesArray = [inputNames];
    }
    if (!Array.isArray(inputNamesArray)) {
      console.log(`NoHands: Skipping ${columnName} - invalid mapping format`);
      continue;
    }

    // Fill each mapped input
    inputNamesArray.forEach(inputName => {
      if (!inputName || inputName.trim() === '') {
        return;
      }

      try {
        // Find input by name attribute
        const input = document.querySelector(`[name="${inputName}"]`);

        if (!input) {
          const error = `Input non trouvé: ${inputName}`;
          console.warn(`NoHands: ${error}`);
          errors.push(error);
          return;
        }

        // Fill based on input type
        const success = fillInputByType(input, value);
        if (success) {
          filledCount++;
          filled.push(`${columnName} → ${inputName}`);
          console.log(`NoHands: Filled ${columnName} (${inputName}) with "${value}"`);
        } else {
          const error = `Échec pour ${columnName} → ${inputName}`;
          console.warn(`NoHands: ${error}`);
          errors.push(error);
        }

      } catch (error) {
        const errorMsg = `Erreur pour ${columnName} → ${inputName}: ${error.message}`;
        console.error(`NoHands: ${errorMsg}`, error);
        errors.push(errorMsg);
      }
    });
  }

  const result = {
    success: filledCount > 0,
    filledCount: filledCount,
    filled: filled,
    errors: errors.length > 0 ? errors : null,
    error: errors.length > 0 ? errors.slice(0, 3).join(', ') : null
  };

  console.log('NoHands: Fill complete', result);
  return result;
}

/**
 * Fill form fields from custom name -> value pairs (no column mapping)
 * @param {Object.<string, string>} customFields - Input name to value
 * @returns {Object} - Result with filledCount, filled, errors
 */
function fillCustomFields(customFields) {
  let filledCount = 0;
  const errors = [];
  const filled = [];

  for (const [inputName, value] of Object.entries(customFields)) {
    if (!inputName || inputName.trim() === '') continue;

    try {
      const input = document.querySelector(`[name="${inputName}"]`);
      if (!input) {
        errors.push(`Input non trouvé: ${inputName}`);
        continue;
      }
      const success = fillInputByType(input, value);
      if (success) {
        filledCount++;
        filled.push(`custom:${inputName}`);
        console.log(`NoHands: Filled custom ${inputName} with "${value}"`);
      } else {
        errors.push(`Échec pour custom → ${inputName}`);
      }
    } catch (error) {
      errors.push(`Erreur custom ${inputName}: ${error.message}`);
    }
  }

  return { filledCount, filled, errors };
}

/**
 * Fill input based on its type
 * @param {HTMLElement} input - The input element
 * @param {string} value - The value to fill
 * @returns {boolean} - Success or failure
 */
function fillInputByType(input, value) {
  const tagName = input.tagName.toLowerCase();
  const type = input.type ? input.type.toLowerCase() : 'text';

  console.log(`NoHands: Filling ${tagName} input of type "${type}" with value "${value}"`);

  // Text inputs (text, email, tel, number, url, search, password)
  if (tagName === 'input' && ['text', 'email', 'tel', 'number', 'url', 'search', 'password'].includes(type)) {
    input.value = value;
    triggerInputEvents(input);
    return true;
  }

  // Textarea
  if (tagName === 'textarea') {
    input.value = value;
    triggerInputEvents(input);
    return true;
  }

  // Select dropdown
  if (tagName === 'select') {
    const valueLower = value.toString().toLowerCase().trim();

    // Try multiple matching strategies
    let option = null;

    // 1. Exact match on value
    option = Array.from(input.options).find(opt => opt.value === value);

    // 2. Exact match on text
    if (!option) {
      option = Array.from(input.options).find(opt => opt.text === value);
    }

    // 3. Case-insensitive match on value
    if (!option) {
      option = Array.from(input.options).find(opt =>
        opt.value.toLowerCase().trim() === valueLower
      );
    }

    // 4. Case-insensitive match on text
    if (!option) {
      option = Array.from(input.options).find(opt =>
        opt.text.toLowerCase().trim() === valueLower
      );
    }

    // 5. Partial match on text (contains)
    if (!option) {
      option = Array.from(input.options).find(opt =>
        opt.text.toLowerCase().includes(valueLower) || valueLower.includes(opt.text.toLowerCase())
      );
    }

    // 6. Partial match on value (contains)
    if (!option) {
      option = Array.from(input.options).find(opt =>
        opt.value.toLowerCase().includes(valueLower) || valueLower.includes(opt.value.toLowerCase())
      );
    }

    if (option) {
      input.value = option.value;
      triggerChangeEvent(input);
      console.log(`NoHands: Select matched "${value}" to option "${option.text}" (value: ${option.value})`);
      return true;
    } else {
      console.warn(`NoHands: No matching option found in select for value "${value}"`);
      console.log(`NoHands: Available options:`, Array.from(input.options).map(opt => `${opt.text} (${opt.value})`));
      return false;
    }
  }

  // Checkbox
  if (tagName === 'input' && type === 'checkbox') {
    // Check if value indicates checked state
    const shouldCheck = ['o', 'oui', 'yes', 'true', '1', 'on', 'checked'].includes(
      value.toString().toLowerCase()
    );
    input.checked = shouldCheck;
    triggerChangeEvent(input);
    return true;
  }

  // Radio buttons
  if (tagName === 'input' && type === 'radio') {
    // Find all radios with the same name
    const radios = document.querySelectorAll(`[name="${input.name}"]`);
    const matchingRadio = Array.from(radios).find(r =>
      r.value === value || r.value.toLowerCase() === value.toLowerCase()
    );

    if (matchingRadio) {
      matchingRadio.checked = true;
      triggerChangeEvent(matchingRadio);
      return true;
    } else {
      console.warn(`NoHands: No matching radio button found for value "${value}"`);
      return false;
    }
  }

  // Date input
  if (tagName === 'input' && type === 'date') {
    // Convert date format if necessary (DD/MM/YYYY → YYYY-MM-DD)
    const convertedDate = convertDateFormat(value);
    if (convertedDate) {
      input.value = convertedDate;
      triggerChangeEvent(input);
      return true;
    } else {
      console.warn(`NoHands: Could not convert date format for "${value}"`);
      return false;
    }
  }

  // Hidden input (may be used by some frameworks)
  if (tagName === 'input' && type === 'hidden') {
    input.value = value;
    triggerChangeEvent(input);
    return true;
  }

  // Unsupported type
  console.warn(`NoHands: Unsupported input type: ${tagName} (${type})`);
  return false;
}

/**
 * Trigger input and change events for frameworks to detect changes
 * @param {HTMLElement} element - The element to trigger events on
 */
function triggerInputEvents(element) {
  // Trigger input event
  element.dispatchEvent(new Event('input', { bubbles: true }));
  // Trigger change event
  element.dispatchEvent(new Event('change', { bubbles: true }));
  // Trigger blur event (some frameworks listen to this)
  element.dispatchEvent(new Event('blur', { bubbles: true }));
}

/**
 * Trigger change event only
 * @param {HTMLElement} element - The element to trigger events on
 */
function triggerChangeEvent(element) {
  element.dispatchEvent(new Event('change', { bubbles: true }));
  element.dispatchEvent(new Event('blur', { bubbles: true }));
}

/**
 * Convert date from DD/MM/YYYY to YYYY-MM-DD format
 * @param {string} dateStr - Date string to convert
 * @returns {string|null} - Converted date or null if invalid
 */
function convertDateFormat(dateStr) {
  // Try DD/MM/YYYY format
  const match1 = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match1) {
    const [, day, month, year] = match1;
    return `${year}-${month}-${day}`;
  }

  // Try D/M/YYYY format
  const match2 = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match2) {
    const [, day, month, year] = match2;
    const paddedDay = day.padStart(2, '0');
    const paddedMonth = month.padStart(2, '0');
    return `${year}-${paddedMonth}-${paddedDay}`;
  }

  // Try YYYY-MM-DD format (already correct)
  const match3 = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match3) {
    return dateStr;
  }

  // Return as-is and let the browser handle it
  return dateStr;
}

/**
 * Show a temporary notification near the input
 * @param {HTMLElement} element - The input element
 * @param {string} inputName - The copied input name
 */
function showCopyNotification(element, inputName) {
  // Create notification element
  const notification = document.createElement('div');
  notification.textContent = `✓ Copié: ${inputName}`;
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: rgba(24, 24, 27, 0.92);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    color: #4ade80;
    padding: 12px 20px;
    border-radius: 12px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    font-size: 14px;
    font-weight: 600;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3), 0 0 0 1px rgba(255, 255, 255, 0.06);
    z-index: 999999;
    animation: slideInRight 0.3s ease;
  `;

  // Add animation keyframes
  if (!document.getElementById('nohands-notification-styles')) {
    const style = document.createElement('style');
    style.id = 'nohands-notification-styles';
    style.textContent = `
      @keyframes slideInRight {
        from {
          transform: translateX(100%);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
      @keyframes slideOutRight {
        from {
          transform: translateX(0);
          opacity: 1;
        }
        to {
          transform: translateX(100%);
          opacity: 0;
        }
      }
    `;
    document.head.appendChild(style);
  }

  // Add to page
  document.body.appendChild(notification);

  // Remove after 2 seconds
  setTimeout(() => {
    notification.style.animation = 'slideOutRight 0.3s ease';
    setTimeout(() => {
      notification.remove();
    }, 300);
  }, 2000);
}

console.log('NoHands: Content script loaded');

} // End of injection guard
