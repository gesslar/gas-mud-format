const doWrap = indent => {
  const selected = getSelectedText();
  const finished = [];

  Logger.log(selected.length);

  selected.forEach(currentParagraph => {
    const workingLines = currentParagraph.split("\r");
    const mappedLines = workingLines.map(line => {
      asciiCheck(line);
      return perform_wrap_function(line, indent);
    });
    finished.push(...mappedLines);
  });
  Logger.log(finished.join("\r"));
  insertText(finished.join("\r"));
  return true;
};

const asciiCheck = txt => {
  for (let c = 0, chars = txt.length; c < chars; c++) {
    if (!validateAscii(txt.charCodeAt(c))) {
      throw new Error(
        `Unidentified ASCII character \"${c}\" in selected text.`
      );
    }
  }
  return true;
};

const validateAscii = char => {
  return char >= 32 && char <= 126;
};

const perform_wrap_function = (txt, indent = false) => {
  const MAX_WIDTH = 79;
  const len = txt.length,
    indent_width = 5;
  let counter = 0,
    position = 0,
    space = 0,
    skip = 0;
  let result = "",
    indent_txt;

  if (len <= MAX_WIDTH) return txt;

  if (indent) indent_txt = " ".repeat(indent_width);
  else indent_txt = "";

  while (counter < len) {
    if (position === counter) {
      // on a new line, so indent and reset skip counter.
      if (indent && counter) {
        skip = indent_width;
        result += indent_txt;
      } else {
        skip = 0;
      }
    }

    if (txt.substr(counter, 1) === " ") space = counter;

    if (skip + counter - position >= MAX_WIDTH) {
      if (counter - space < 15) {
        // line wrap instead of word wrap if no recent spaces.
        Logger.log(
          `POSITION 1: ${position} COUNTER 1: ${counter} SPACE 1: ${space}`
        );
        result += txt.slice(position, space) + "\r";
        position = counter = ++space;
      } else {
        Logger.log(
          `POSITION 2: ${position} COUNTER 2: ${counter} SPACE 2: ${space}`
        );
        result += txt.slice(position, counter) + "\r";
        position = counter;
        space = counter;
        // position = space = counter;
      }
      continue;
    }
    counter++;
  }

  if (position != counter) result += txt.slice(position, counter) + "\r";

  return result;
};

function wrapIt() {
  if (doWrap(false)) return "All wrapped up.";
  throw new Error("Unable to wrap the text.");
}

function iwrapIt() {
  if (doWrap(true)) return "All indent wrapped up.";
  throw new Error("Unable to indent wrap the text.");
}

/**
 * Gets the text the user has selected. If there is no selection,
 * this function displays an error message.
 *
 * @return {Array.<string>} The selected text.
 */
function getSelectedText() {
  const selection = DocumentApp.getActiveDocument().getSelection();

  if (selection) {
    const text = [];
    const selectedElements = selection.getSelectedElements();

    selectedElements.forEach(selectedElement => {
      if (selectedElement.isPartial()) {
        const elementText = selectedElement.getElement().asText();
        const startIndex = selectedElement.getStartOffset();
        const endIndex = selectedElement.getEndOffsetInclusive();

        text.push(elementText.getText().substring(startIndex, endIndex + 1));
      } else {
        const element = selectedElement.getElement();
        if (element.editAsText) {
          const elementText = element.asText().getText();

          if (elementText) {
            text.push(elementText);
          }
        }
      }
    });

    if (!text.length) throw new Error("Please select some text.");
    return text;
  } else {
    throw new Error("Please select some text.");
  }
}

/**
 * Replaces the text of the current selection with the provided text, or
 * inserts text at the current cursor location. (There will always be either
 * a selection or a cursor.) If multiple elements are selected, only inserts the
 * translated text in the first element that can contain text and removes the
 * other elements.
 *
 * @param {string} newText The text with which to replace the current selection.
 */
function insertText(newText) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var replaced = false;
    var elements = selection.getSelectedElements();
    if (
      elements.length === 1 &&
      elements[0].getElement().getType() ===
        DocumentApp.ElementType.INLINE_IMAGE
    ) {
      throw new Error("Can't insert text into an image.");
    }
    for (var i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        element.deleteText(startIndex, endIndex);
        if (!replaced) {
          element.insertText(startIndex, newText);
          element.setFontFamily("Roboto Mono");
          element.setFontSize(8);
          replaced = true;
        } else {
          // This block handles a selection that ends with a partial element. We
          // want to copy this partial text to the previous element so we don't
          // have a line-break before the last partial.
          var parent = element.getParent();
          var remainingText = element.getText().substring(endIndex + 1);
          parent
            .getPreviousSibling()
            .asText()
            .appendText(remainingText);
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just remove the text within the last paragraph instead.
          if (parent.getNextSibling()) {
            parent.removeFromParent();
          } else {
            element.removeFromParent();
          }
        }
      } else {
        var element = elements[i].getElement();
        if (!replaced && element.editAsText) {
          // Only translate elements that can be edited as text, removing other
          // elements.
          element.clear();
          element.asText().setText(newText);
          element.setFontFamily("Roboto Mono");
          element.setFontSize(8);
          replaced = true;
        } else {
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just clear the element.
          if (element.getNextSibling()) {
            element.removeFromParent();
          } else {
            element.clear();
          }
        }
      }
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var surroundingText = cursor.getSurroundingText().getText();
    var surroundingTextOffset = cursor.getSurroundingTextOffset();

    // If the cursor follows or preceds a non-space character, insert a space
    // between the character and the translation. Otherwise, just insert the
    // translation.
    if (surroundingTextOffset > 0) {
      if (surroundingText.charAt(surroundingTextOffset - 1) != " ") {
        newText = " " + newText;
      }
    }
    if (surroundingTextOffset < surroundingText.length) {
      if (surroundingText.charAt(surroundingTextOffset) != " ") {
        newText += " ";
      }
    }
    cursor.insertText(newText);
  }
}
