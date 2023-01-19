// Copyright (c) 2015, Jörg Knobloch. All rights reserved.

/* global ExtensionCommon */

const { Services } = ChromeUtils.import("resource://gre/modules/Services.jsm");
const { MailServices } = ChromeUtils.import("resource:///modules/MailServices.jsm");
var { ExtensionSupport } = ChromeUtils.import("resource:///modules/ExtensionSupport.jsm");
var { AppConstants } = ChromeUtils.import("resource://gre/modules/AppConstants.jsm");

const EXTENSION_NAME = "JorgK@dictionaryforrecipient";

const nsIAbDirectory = Ci.nsIAbDirectory;

const verbose = 0;

function getCardForEmail(emailAddress) {
  // copied from msgHdrViewOverlay.js
  var books = MailServices.ab.directories;
  var result = { book: null, card: null };
  for (let ab of books) {
    try {
      var card = ab.cardForEmailAddress(emailAddress);
      if (card) {
        result.book = ab;
        result.card = card;
        break;
      }
    } catch (ex) {
      if (verbose) console.error(`Error (${ex.message}) fetching card from address book ${ab.dirName} for ${emailAddress}`);
    }
  }
  if (verbose && !result.card) console.log(`No address book entry for ${emailAddress}`);
  if (verbose && result.book) console.log(`${emailAddress} served from address book ${result.book.dirName}`);
  return result;
}

function getDictForAddress(address) {
  var cardDetails = getCardForEmail(address);
  if (!cardDetails.card) return null;
  var dict = cardDetails.card.getProperty("Custom4", "");
  if (!dict) {
    try {
      dict = cardDetails.card.vCardProperties.getFirstValue("x-custom4");
    } catch (ex) {}
  }
  return dict;
}

function setDictionary(window) {
  if (verbose) console.log("Setting Dictionary");

  // Get the first recipient.
  var fields = Components.classes["@mozilla.org/messengercompose/composefields;1"]
    .createInstance(Components.interfaces.nsIMsgCompFields);
  window.Recipients2CompFields(fields);
  var recipients = { length: 0 };

  var fields_content = fields.to;
  if (fields_content) {
    recipients = fields.splitRecipients(fields_content, true, {});
  }

  var firstRecipient = null;
  if (recipients.length > 0) {
    firstRecipient = recipients[0].toString();
  }

  var existingRecipient = window.dictionaryForRecipient;
  if (verbose) console.log(`Found existing recipient |${existingRecipient}|`);

  var infoText = null;
  var storeRecipient = null;
  if (firstRecipient) {
    if (firstRecipient == existingRecipient) {
      // We only set the dictionary once, as long as the recipient doesn't change.
      // This allows the user to change the dictionary, if they want to use a different
      // language as an exception. We don't want to reset the user choice.
      if (verbose) infoText = `Not setting dictionary for unchanged recipient ${firstRecipient}`;
    } else {
      var dict = getDictForAddress(firstRecipient);
      if (dict) {
        infoText = `${firstRecipient} -> ${dict}`;
        storeRecipient = firstRecipient;
        if (parseInt(AppConstants.MOZ_APP_VERSION, 10) >= 102) {
          // From 102 onwards we can have multiple languages.
          // Switch them all off before setting our own.
          window.gActiveDictionaries = new Set();
        }
        var changeEvent = { target: { value: dict }, stopPropagation() {} };
        window.ChangeLanguage(changeEvent);
      } else {
        if (verbose) infoText = `Setting dictionary: ${firstRecipient} has no dictionary defined`;
        storeRecipient = "-";
      }
    }
  } else {
    if (verbose) infoText = "Setting dictionary: No recipient specified";
    storeRecipient = "-";
  }
  if (verbose) console.log(infoText);

  if (storeRecipient && storeRecipient != existingRecipient) {
    window.dictionaryForRecipient = storeRecipient;
    if (verbose) console.log(`Storing recipient |${storeRecipient}|`);
  }

  if (infoText) {
    var statusText = window.document.getElementById("statusText");
    // statusText.label = infoText;  // TB 68
    statusText.setAttribute("value", infoText);  // TB 70 and later, bug 1577659.
    statusText.textContent = infoText;  // Statusbar as HTML.
    window.setTimeout(() => {
      // statusText.label = "";  // TB 68
      statusText.setAttribute("value", "");  // TB 70 and later, bug 1577659.
      statusText.textContent = "";  // Statusbar as HTML.
    }, 2000);
  }
}

function setListener1(window, target, on) {
  // Use closure to carry 'window' and 'on' into the callback function.
  var listener1 = function (evt) {
    if (verbose) console.log(`received event ${on}`);
    // We can only set new dictionaries after the initial spell check
    // triggered by switching on the inline spell checker is finished.
    if (!window.gSpellCheckingEnabled) {
      setDictionary(window);
    } else {
      (function setDict() {
        if (verbose) console.log(`We can switch dictionaries now: ${window.checkerReadyObserver_Dict_for_recipient_add_on.isReady()}`);
        if (window.checkerReadyObserver_Dict_for_recipient_add_on.isReady()) {
          window.checkerReadyObserver_Dict_for_recipient_add_on.removeObserver();
          // In theory, a furthter timeout shouldn't be necessary here any more :-(
          // Observed behaviour is that it still fails one out of seven without the timeout.
          window.setTimeout(() => {
            setDictionary(window);
          }, 500);
        } else {
          window.setTimeout(setDict, 100);
        }
      }());
    }
  };
  target.addEventListener(on, listener1);
}

/**
 * We prepare the compose window by attaching our listeners and resetting our properties.
 */
function PrepareComposeWindow(window) {
  if (verbose) console.log("Preparing compose window");
  window.checkerReadyObserver_Dict_for_recipient_add_on = {
    _topic: "inlineSpellChecker-spellCheck-ended",
    _isReady: false,

    observe(aSubject, aTopic, aData) {
      if (aTopic != this._topic) {
        return;
      }
      this._isReady = true;
      if (verbose) console.log(`checkerReadyObserver_Dict_for_recipient_add_on.observe triggered: ${this._isReady}`);
    },

    _isAdded: false,

    addObserver() {
      if (this._isAdded) {
        return;
      }

      Services.obs.addObserver(this, this._topic);
      this._isAdded = true;
    },

    removeObserver() {
      if (!this._isAdded) {
        return;
      }

      Services.obs.removeObserver(this, this._topic);
      this._isAdded = false;
      // this._isReady = false;
    },

    isReady() {
      return this._isReady;
    },
  };
  window.checkerReadyObserver_Dict_for_recipient_add_on.addObserver();

  window.dictionaryForRecipient = "-";

  // If these events arrive, we need to derive the dictionary again.
  setListener1(window, window.document.getElementById("msgSubject"), "focus");
  setListener1(window, window, "blur");
}

function CleanupComposeWindow(window) {
  if (verbose) console.log("Cleaning up compose window");
  window.checkerReadyObserver_Dict_for_recipient_add_on.removeObserver();
  delete window.checkerReadyObserver_Dict_for_recipient_add_on;
}

// Implements the functions defined in the experiments section of schema.json.
var DictionaryForRecipient = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    return {
      DictionaryForRecipient: {
        addComposeWindowListener(dummy) {
          // Adds a listener to detect new compose windows.
          if (verbose) console.log("DictionaryForRecipient: addComposeWindowListener");

          ExtensionSupport.registerWindowListener(EXTENSION_NAME, {
            chromeURLs: ["chrome://messenger/content/messengercompose/messengercompose.xul",
              "chrome://messenger/content/messengercompose/messengercompose.xhtml"],
            onLoadWindow: PrepareComposeWindow,
            onUnloadWindow: CleanupComposeWindow,
          });
        },
      },
    };
  }

  onShutdown(isAppShutdown) {
    ExtensionSupport.unregisterWindowListener(EXTENSION_NAME);
  }
};
