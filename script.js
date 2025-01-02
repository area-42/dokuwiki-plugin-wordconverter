/**
 * DokuWiki Word Converter Plugin
 */
if (typeof window.toolbar !== 'undefined') {
  // Register Toolbar Button
  window.toolbar.push({
    type:  'WordImport',
    title: 'Import from Word',
    icon:  '../../plugins/wordconverter/clipboard.png'
  });

  function addBtnActionWordImport($btn, props, edid) {
    $btn.click(async function() {
      try {
        // Get current selection/cursor position
        let selection      = DWgetSelection(jQuery('#wiki__text')[0]);
        let clipboardItems = await navigator.clipboard.read();

        for (let clipboardItem of clipboardItems) {
          for (let type of clipboardItem.types) {
            if (type === "text/html") {
              let blob      = await clipboardItem.getType(type);
              let html      = await blob.text();
              let converted = convertWordToWiki(html);

              // Insert at cursor position
              pasteText(selection, converted);
              return;
            }
          }
        }

        alert('No HTML format found in the clipboard.');
      } catch (err) {
        alert('Error accessing the clipboard: ' + err.message);
        console.error(err);
      }
    });

    return 'wordimport';
  }

  function convertWordToWiki(html) {
    let parser    = new DOMParser();
    let doc       = parser.parseFromString(html, 'text/html');
    let converted = convertNodeToDokuWiki(doc.body);

    return converted.replace(/\n{3,}/g, '\n\n').trim();
  }

  function convertNodeToDokuWiki(node, nestingLevel = 0) {
    if (!node) {
      return '';
    }

    // Return Text Node directly
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent.replace(/\s+/g, ' ');
    }

    // Calculate indentation for nested lists
    let indent = '  '.repeat(nestingLevel);

    switch (node.tagName?.toLowerCase()) {
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        let headerLevel = parseInt(node.tagName.charAt(1));
        let equals      = '='.repeat(7 - headerLevel);

        return `\n${equals} ${processInlineContent(node)} ${equals}\n`;

      case 'p':
        let content = processInlineContent(node);

        if (!content.trim()) {
          return '\n';
        }

        return `\n${content}\n`;

      case 'br':
        return '\n';

      case 'strong':
      case 'b':
        return `**${processInlineContent(node)}**`;

      case 'em':
      case 'i':
        return `//${processInlineContent(node)}//`;

      case 'u':
        return `__${processInlineContent(node)}__`;

      case 'a':
        let href = node.getAttribute('href');

        if (href) {
          return `[[${href}|${processInlineContent(node)}]]`;
        }

        return processInlineContent(node);

      case 'ul':
        return formatList(node, '*', nestingLevel);

      case 'ol':
        return formatList(node, '-', nestingLevel);

      case 'pre':
      case 'code':
        return `\n<code>\n${processInlineContent(node)}\n</code>\n`;

      case 'table':
        return convertTable(node);

      default:
        return Array.from(node.childNodes)
          .map(child => convertNodeToDokuWiki(child, nestingLevel))
          .join('');
    }
  }

  function formatList(node, marker, nestingLevel) {
    // Iterate over the list elements
    let items = Array.from(node.children).map(li => {
      let indent = '  '.repeat(nestingLevel + 1); // Two spaces even for the top level

      // Check for nested lists
      let nestedLists = Array.from(li.children).filter(child => child.tagName?.toLowerCase() === 'ul' || child.tagName?.toLowerCase() === 'ol');

      // Process the main content of the list element
      let mainContent = Array.from(li.childNodes)
        .filter(node => node.nodeType === Node.TEXT_NODE || (node.nodeType === Node.ELEMENT_NODE && node.tagName?.toLowerCase() !== 'ul' && node.tagName?.toLowerCase() !== 'ol'))
        .map(node => convertNodeToDokuWiki(node, nestingLevel))
        .join('').trim();

      // Process nested lists
      let nestedContent = nestedLists
        .map(list => formatList(list, marker, nestingLevel + 1)) // Call recursively
        .join('\n');

      return `${indent}${marker} ${mainContent}${nestedContent? '\n' + nestedContent : ''}`;
    }).join('\n');

    // Insert blank lines for the top level only
    return ((nestingLevel === 0)? '\n' : '') + items;
  }

  function processInlineContent(node) {
    return Array.from(node.childNodes)
      .map(child => convertNodeToDokuWiki(child))
      .join('')
      .trim();
  }

  function convertTable(table) {
    let result  = '\n';
    let headers = Array.from(table.querySelectorAll('th'));

    if (headers.length) {
      result += '^ ' + headers.map(th => processInlineContent(th)).join(' ^ ') + ' ^\n';
    }

    let rows = Array.from(table.querySelectorAll('tr'));

    for (let row of rows) {
      let cells = Array.from(row.querySelectorAll('td'));

      if (cells.length) {
        result += '| ' + cells.map(td => processInlineContent(td)).join(' | ') + ' |\n';
      }
    }

    return result;
  }

  // Hide functionality for comment section
  jQuery(document).ready(function() {
    jQuery('#discussion__comment_toolbar').children('[aria-controls="wordimport"]').hide();
  });
}
