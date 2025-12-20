// Red Ink Browser Extension
// Copyright by David Rosenthal, david.rosenthal@vischer.com
// May only be used with permission
// 30.11.2025

/****************************************
 *  Red Ink Service Worker / Background
 ****************************************/

// Set the extension API depending on the type of browser the extension is installed on
const extensionAPI = typeof browser !== 'undefined' ? browser : chrome;

// Log when the service worker starts
console.log("Red Ink Browser Extension (Version 30.11.2025) service worker loaded successfully!");

// Create the context menu hierarchy when the extension is installed
extensionAPI.runtime.onInstalled.addListener(() => {
    console.log("Red Ink Browser Extension installed. Creating context menus...");

    // Parent menu
    extensionAPI.contextMenus.create({
        id: "redink_root",
        title: "Red Ink",
        contexts: ["all"], // Show the menu on all contexts
    });

    // Submenu items
    extensionAPI.contextMenus.create({
        id: "redink_translate",
        title: "Translate",
        parentId: "redink_root",
        contexts: ["selection"], // Only valid if there's text selected
    });

    extensionAPI.contextMenus.create({
        id: "redink_correct",
        title: "Correct",
        parentId: "redink_root",
        contexts: ["selection"],
    });

    extensionAPI.contextMenus.create({
        id: "redink_freestyle",
        title: "Freestyle",
        parentId: "redink_root",
        contexts: ["all"], // Always enabled
    });

    extensionAPI.contextMenus.create({
        id: "redink_sendtoword",
        title: "Send to Word",
        parentId: "redink_root",
        contexts: ["selection"],
    });

    extensionAPI.contextMenus.create({
        id: "redink_sendtooutlook",
        title: "Send to Outlook",
        parentId: "redink_root",
        contexts: ["selection"],
    });

    console.log("Context menus created.");
});

// Listener for menu clicks
extensionAPI.contextMenus.onClicked.addListener((info, tab) => {
    const command = info.menuItemId;
    const selectedText = info.selectionText || "";

    // Prevent actions for commands other than "Freestyle" when no text is selected
    if (!selectedText && command !== "redink_freestyle") {
        console.log(`Command "${command}" requires a text selection.`);
        return;
    }

    switch (command) {
        case "redink_translate":
            handleTranslate(info, tab);
            break;
        case "redink_correct":
            handleCorrect(info, tab);
            break;
        case "redink_freestyle":
            handleFreestyle(info, tab);
            break;
        case "redink_sendtoword":
            handleSendToWord(info, tab);
            break;
        case "redink_sendtooutlook":
            handleSendToOutlook(info, tab);
            break;
        default:
            console.error("Unknown Red Ink command: " + command);
    }
});



/***************************************************
 *                 COMMAND HANDLERS
 **************************************************/

// 1. Translate
function handleTranslate(info, tab) {
    // No prompt needed per instructions. Instruction = ""
    sendRequestToLocalHost({
        command: info.menuItemId,
        instruction: "",
        text: info.selectionText || "",
        tab: tab
    });
}

// 2. Correct
function handleCorrect(info, tab) {
    // No prompt needed per instructions. Instruction = ""
    sendRequestToLocalHost({
        command: info.menuItemId,
        instruction: "",
        text: info.selectionText || "",
        tab: tab
    });
}

// 3. Freestyle
function handleFreestyle(info, tab) {
    // No prompt needed per instructions. Instruction = ""
    sendRequestToLocalHost({
        command: info.menuItemId,
        instruction: "",
        text: info.selectionText || "",
        tab: tab
    });
}

// 4. Send to Word
function handleSendToWord(info, tab) {
    // No prompt needed per instructions. Instruction = ""
    sendRequestToLocalHost({
        command: info.menuItemId,
        instruction: "",
        text: info.selectionText || "",
        tab: tab
    });
}

// 5. Send to Outlook
function handleSendToOutlook(info, tab) {
    // No prompt needed per instructions. Instruction = ""
    sendRequestToLocalHost({
        command: info.menuItemId,
        instruction: "",
        text: info.selectionText || "",
        tab: tab
    });
}

/***************************************************
 *                  NETWORK LOGIC
 **************************************************/

function sendRequestToLocalHost({ command, instruction, text, tab }) {
    const requestBody = {
        URL: tab.url || "",
        Command: command,
        Instruction: instruction,
        Text: text
    };

    let port = "12333";
    if (command === "redink_sendtoword") {
        port = "12334";
    }

    const localEndpoint = `http://127.0.0.1:${port}/redink`;

    fetch(localEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
    })
        .then(response => {
            if (!response.ok) {
                throw new Error(`Server responded with status ${response.status}`);
            }
            if (command === "redink_sendtoword" || command === "redink_sendtooutlook") {
                return null; // No text insertion expected for these commands
            }
            return response.text();
        })
        .then(processedText => {
            if (processedText === null || processedText === "") {
                // Do not replace or insert anything if processedText is empty or null
                console.log("No text returned from the remote application. Skipping insertion.");
                return;
            }
            console.log("Received processed text:", processedText);
            return extensionAPI.scripting.executeScript({
                target: { tabId: tab.id },
                func: replaceSelectedText,
                args: [processedText]
            });
        })
        .then(() => {
            // Best-effort: bring tab and window to foreground (requires tabs/windows permissions)
            if (tab && typeof tab.id === "number") {
                try { extensionAPI.tabs.update(tab.id, { active: true }); } catch (e) { console.warn("tabs.update failed:", e); }
            }
            if (tab && typeof tab.windowId === "number" && extensionAPI.windows?.update) {
                try { extensionAPI.windows.update(tab.windowId, { focused: true }); } catch (e) { console.warn("windows.update failed:", e); }
            }
        })
        .catch(error => {
            console.error(`Error communicating with the Red Ink host: ${error.message}`);
        });
}

/***************************************************
 *            TEXT INSERTION HELPER
 **************************************************/

// Replaces whatever is selected in the DOM with `newText`
function replaceSelectedText(newText) {
    if (newText == null) return;

    const active = document.activeElement;

    // Case 1: input or textarea
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA")) {
        const start = active.selectionStart ?? 0;
        const end = active.selectionEnd ?? start;

        if (typeof active.setRangeText === "function") {
            active.setRangeText(newText, start, end, "end");
        } else {
            const val = active.value ?? "";
            active.value = val.slice(0, start) + newText + val.slice(end);
            const caret = start + newText.length;
            active.setSelectionRange?.(caret, caret);
        }

        active.dispatchEvent(new Event("input", { bubbles: true }));
        active.dispatchEvent(new Event("change", { bubbles: true }));
        return;
    }

    // Case 2: any DOM selection (editable or not)
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;

    const range = selection.getRangeAt(0);
    // Replace the selected DOM contents with a plain text node
    range.deleteContents();
    const textNode = document.createTextNode(newText);
    range.insertNode(textNode);

    // Move caret to the end of the inserted text and restore selection
    range.setStartAfter(textNode);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
}



