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
    $btn.click(function() {
      // Move async logic into a separate function
      handleWordImport().catch(err => {
        alert('Error: ' + err.message);
        console.error(err);
      });
    });

    return 'wordimport';
  }

  // Separate async logic
  async function handleWordImport() {
    let selection      = DWgetSelection(jQuery('#wiki__text')[0]);
    let clipboardItems = await navigator.clipboard.read();

    for (let clipboardItem of clipboardItems) {
      for (let type of clipboardItem.types) {
        if (type === "text/html") {
          let blob      = await clipboardItem.getType(type);
          let html      = await blob.text();
          let converted = await convertWordToWiki(html);
          
          // Insert at cursor position
          pasteText(selection, converted);
          return;
        }
      }
    }

    alert('No HTML format found in the clipboard.');
  }

  async function convertWordToWiki(html) {
    let parser    = new DOMParser();
    let doc       = parser.parseFromString(html, 'text/html');
    let converted = await convertNodeToDokuWiki(doc.body);

    return converted.replace(/\n{3,}/g, '\n\n').trim();
  }

  async function convertNodeToDokuWiki(node, nestingLevel = 0) {
    if (!node) {
      return '';
    }

    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent.replace(/\s+/g, ' ');
    }

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
        let content     = await processInlineContent(node);

        return `\n${equals} ${content} ${equals}\n`;

      case 'p':
        let pContent = await processInlineContent(node);

        if (!pContent.trim()) {
          return '\n';
        }

        return `\n${pContent}\n`;

      case 'br':
        return '\n';

      case 'strong':
      case 'b':
        return `**${await processInlineContent(node)}**`;

      case 'em':
      case 'i':
        return `//${await processInlineContent(node)}//`;

      case 'u':
        return `__${await processInlineContent(node)}__`;

      case 'a':
        let href = node.getAttribute('href');

        if (href) {
          return `[[${href}|${await processInlineContent(node)}]]`;
        }

        return await processInlineContent(node);

      case 'ul':
        return formatList(node, '*', nestingLevel);

      case 'ol':
        return formatList(node, '-', nestingLevel);

      case 'pre':
      case 'code':
        return `\n<code>\n${await processInlineContent(node)}\n</code>\n`;

      case 'table':
        return await convertTable(node);

      case 'img':
        return await handleImage(node);

      case 'del':
        return `<del>${await processInlineContent(node)}</del>`;

      case 'sub':
        return `<sub>${await processInlineContent(node)}</sub>`;

      case 'sup':
        return `<sup>${await processInlineContent(node)}</sup>`;

      case 'blockquote':
        return `> ${await processInlineContent(node)}`;

      case 'hr':
        return `\n----\n`;

      case 'dl':
        return Array.from(node.children).map(child => {
          if (child.tagName?.toLowerCase() === 'dt') {
            return `\n${processInlineContent(child)}`;
          } else if (child.tagName?.toLowerCase() === 'dd') {
            return `: ${processInlineContent(child)}`;
          }
        }).join('\n');

      case 'kbd':
        return `<kbd>${await processInlineContent(node)}</kbd>`;

      default:
        let results = await Promise.all(
          Array.from(node.childNodes).map(child => 
            convertNodeToDokuWiki(child, nestingLevel)
          )
        );

        return results.join('');
    }
  }

  async function formatList(node, marker, nestingLevel) {
    let itemsPromises = Array.from(node.children).map(async li => {
      let indent      = '  '.repeat(nestingLevel + 1);
      let nestedLists = Array.from(li.children).filter(child => child.tagName?.toLowerCase() === 'ul' || child.tagName?.toLowerCase() === 'ol');

      // Wait for the main content to be processed
      let mainContent = await Promise.all(
        Array.from(li.childNodes)
          .filter(node => node.nodeType === Node.TEXT_NODE || (node.nodeType === Node.ELEMENT_NODE && node.tagName?.toLowerCase() !== 'ul' && node.tagName?.toLowerCase() !== 'ol'))
          .map(node => convertNodeToDokuWiki(node, nestingLevel))
      );

      // Wait for the nested lists to be processed
      let nestedContent = await Promise.all(
        nestedLists.map(list => formatList(list, marker, nestingLevel + 1))
      );

      return `${indent}${marker} ${mainContent.join('').trim()}${nestedContent.length? '\n' + nestedContent.join('\n') : ''}`;
    });

    // Wait for all list entries
    let items = await Promise.all(itemsPromises);
    return ((nestingLevel === 0)? '\n' : '') + items.join('\n');
  }

  async function processInlineContent(node) {
    let results = await Promise.all(
      Array.from(node.childNodes).map(child => 
        convertNodeToDokuWiki(child)
      )
    );

    return results.join('').trim();
  }
  
  async function convertTable(table) {
    let result  = '\n';
    let headers = Array.from(table.querySelectorAll('th'));

    if (headers.length) {
      let headerContents = await Promise.all(
        headers.map(th => processInlineContent(th))
      );

      result += '^ ' + headerContents.join(' ^ ') + ' ^\n';
    }

    let rows = Array.from(table.querySelectorAll('tr'));

    for (let row of rows) {
      let cells = Array.from(row.querySelectorAll('td'));

      if (cells.length) {
        let cellContents = await Promise.all(
          cells.map(td => processInlineContent(td))
        );

        result += '| ' + cellContents.join(' | ') + ' |\n';
      }
    }

    return result;
  }

  async function handleImage(imgNode) {
    try {
      let imgData;
      let fileName;
      let mimeType;

      console.log('Starting image processing:', {
        nodeName:   imgNode.nodeName,
        src:        imgNode.src,
        attributes: Object.fromEntries(
          Array.from(imgNode.attributes || []).map(attr => [attr.name, attr.value])
        )
      });

      if (imgNode.src?.startsWith('file:///')) {
        console.log('Found LibreOffice temporary file path, trying clipboard...');
        
        try {
          const clipboardItems = await navigator.clipboard.read();

          console.log('Available clipboard formats:', 
            clipboardItems.map(item => Array.from(item.types)).flat()
          );

          // Search for image data in the clipboard
          for (const clipboardItem of clipboardItems) {
            const imageTypes = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp'];

            for (const type of imageTypes) {
              if (clipboardItem.types.includes(type)) {
                console.log(`Found ${type} in clipboard`);
                const blob = await clipboardItem.getType(type);

                imgData  = blob;
                mimeType = type;
                fileName = `imported_${Date.now()}.${type.split('/')[1]}`;
                
                break;
              }
            }

            if (imgData) {
              break;
            }

            // If no direct image data is available, check HTML
            if (!imgData && clipboardItem.types.includes('text/html')) {
                console.log('Checking HTML content for embedded images');

                const htmlBlob    = await clipboardItem.getType('text/html');
                const html        = await htmlBlob.text();
                const base64Match = html.match(/data:image\/[^;]+;base64,([^"']+)/);

                if (base64Match) {
                  console.log('Found base64 image in HTML content');

                  const fullDataUrl = base64Match[0];
                  const formatMatch = fullDataUrl.match(/data:image\/([^;]+);/);

                  if (formatMatch) {
                    mimeType = `image/${formatMatch[1]}`;
                    const base64Data = base64Match[1];

                    imgData  = base64ToBlob(base64Data, mimeType);
                    fileName = `imported_${Date.now()}.${formatMatch[1]}`;
                  }
                }
                
                // Search for XML-based LibreOffice image data
                if (!imgData) {
                  const xmlMatch = html.match(/<draw:image[^>]*>([^<]+)<\/draw:image>/);

                  if (xmlMatch) {
                    console.log('Found LibreOffice XML image data');

                    mimeType = 'image/png';  // LibreOffice typically uses PNG
                    imgData  = base64ToBlob(xmlMatch[1], mimeType);
                    fileName = `imported_${Date.now()}.png`;
                  }
                }
            }
          }
        } catch (err) {
          console.error('Clipboard operation failed:', err);
        }
      }
      // Base64 treatment...
      else if (imgNode.src?.startsWith('data:')) {
        console.log('Processing direct base64 image');

        let matches = imgNode.src.match(/^data:([A-Za-z-+/]+);base64,(.+)$/);

        if (matches && matches.length === 3) {
          mimeType = matches[1];
          imgData  = base64ToBlob(matches[2], mimeType);
          fileName = `imported_${Date.now()}.${mimeType.split('/')[1]}`;
        }
      }

      if (imgData) {
        console.log('Image data found, uploading...', {fileName, mimeType});

        let formData = new FormData();
        formData.append('qqfile', imgData, fileName);
        formData.append('call', 'upload');
        formData.append('ow', 'true');

        let uploadResponse = await fetch(DOKU_BASE + 'lib/exe/ajax.php', {
          method:  'POST',
          body:    formData,
          headers: {
            'X-Requested-With': 'XMLHttpRequest'
          }
        });

        let result = await uploadResponse.json();
        console.log('Upload result:', result);

        if (!uploadResponse.ok || result.error) {
          throw new Error(result.error || 'Upload failed');
        }

        let width      = imgNode.width  || '';
        let height     = imgNode.height || '';
        let dimensions = '';

        if (width || height) {
          dimensions = '?' + (width || '') + (height? 'x' + height : '');
        }

        let alt   = imgNode.alt   || imgNode.name || '';
        let title = imgNode.title || imgNode.name || '';
        
        return `\n{{${fileName}${dimensions}|${title}|${alt}}}\n`;
      }

      return `\n// Image could not be imported: No valid image format found. //\n`;
    } catch (err) {
      return `\n// Error during image import: ${err.message} //\n`;
    }
  }

  function base64ToBlob(base64Data, mimeType) {
    let binaryStr = atob(base64Data);
    let bytes     = new Uint8Array(binaryStr.length);

    for (let i = 0; i < binaryStr.length; i++) {
      bytes[i] = binaryStr.charCodeAt(i);
    }
    
    return new Blob([bytes], {
      type: mimeType
    });
  }

  // Hide functionality for comment section
  jQuery(document).ready(function() {
    jQuery('#discussion__comment_toolbar').children('[aria-controls="wordimport"]').hide();
  });
}