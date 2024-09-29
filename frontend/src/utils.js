import JSZip from "jszip";
import * as cheerio from "cheerio";
import { uploadImage } from "./services/api";

export async function extractDocxContent(file) {
  try {
    // Load the DOCX file (which is a zip) using JSZip
    const zip = await JSZip.loadAsync(file);

    // Extract XML files from the zip
    const documentXml = await zip.file("word/document.xml").async("text");
    const stylesXml = await zip.file("word/styles.xml").async("text");
    const numberingXml = await zip.file("word/numbering.xml").async("text");

    // Load the XML content with Cheerio
    const $doc = cheerio.load(documentXml, { xmlMode: true });
    const $styles = cheerio.load(stylesXml, { xmlMode: true });
    const $numbering = cheerio.load(numberingXml, { xmlMode: true });

    // Build numbering map from numbering.xml
    const numberingMap = buildNumberingMap($numbering);

    const styleMap = {};
    $styles("w\\:style").each((_, style) => {
      const styleId = $styles(style).attr("w:styleId");
      const styleType = $styles(style).attr("w:type");

      if (styleId && styleType) {
        styleMap[styleId] = {
          type: styleType,
          name: $styles(style).find("w\\:name").attr("w:val"),
          basedOn: $styles(style).find("w\\:basedOn").attr("w:val"),
          runProperties: extractRunStyles($styles(style).find("w\\:rPr")),
          paragraphProperties: extractParagraphStyles(
            $styles(style).find("w\\:pPr")
          ),
        };
      }
    });

    // Resolve inherited styles based on the 'basedOn' attribute
    Object.keys(styleMap).forEach((styleId) => {
      const style = styleMap[styleId];
      if (style.basedOn) {
        const inheritedStyle = styleMap[style.basedOn];
        if (inheritedStyle) {
          style.runProperties = {
            ...inheritedStyle.runProperties,
            ...style.runProperties,
          };
          style.paragraphProperties = {
            ...inheritedStyle.paragraphProperties,
            ...style.paragraphProperties,
          };
        }
      }
    });

    function extractRunStyles(rPr) {
      const styles = {};

      if (rPr.find("w\\:b").length > 0) styles.bold = true;
      if (rPr.find("w\\:i").length > 0) styles.italic = true;
      if (rPr.find("w\\:u").length > 0) styles.underline = true;
      if (rPr.find("w\\:strike").length > 0) styles.strikeThrough = true;

      const color = rPr.find("w\\:color").attr("w:val");
      if (color) styles.color = color;

      const fontSize = rPr.find("w\\:sz").attr("w:val");
      if (fontSize) styles.fontSize = fontSize;

      const font = rPr.find("w\\:rFonts").attr("w:ascii");
      if (font) styles.font = font;

      const backgroundColor = rPr.find("w\\:shd").attr("w:fill");
      if (backgroundColor) styles.backgroundColor = backgroundColor;

      const highlight = rPr.find("w\\:highlight").attr("w:val");
      if (highlight) styles.highlight = highlight;

      return styles;
    }

    function extractParagraphStyles(pPr) {
      const styles = {};

      const alignment = pPr.find("w\\:jc").attr("w:val");
      if (alignment) styles.alignment = alignment;

      const spacingBefore = pPr.find("w\\:spacing").attr("w:before");
      if (spacingBefore) styles.spacingBefore = spacingBefore;

      const spacingAfter = pPr.find("w\\:spacing").attr("w:after");
      if (spacingAfter) styles.spacingAfter = spacingAfter;

      const indentLeft = pPr.find("w\\:ind").attr("w:left");
      if (indentLeft) styles.indentLeft = indentLeft;

      const indentRight = pPr.find("w\\:ind").attr("w:right");
      if (indentRight) styles.indentRight = indentRight;

      return styles;
    }

    async function parseElement(element) {
      const children = [];

      for (const child of element.children()) {
        const tag = $doc(child)[0].tagName;

        if (tag === "w:p") {
          const paragraphData = {
            type: "paragraph",
            text: "",
            styles: {},
          };

          const pPr = $doc(child).find("w\\:pPr");
          const pStyleId = pPr.find("w\\:pStyle").attr("w:val");

          if (pStyleId && styleMap[pStyleId]) {
            paragraphData.styles = {
              ...styleMap[pStyleId].paragraphProperties,
              ...styleMap[pStyleId].runProperties,
            };
          }

          if (pPr.length) {
            paragraphData.styles = {
              ...paragraphData.styles,
              ...extractParagraphStyles(pPr),
            };

            const numPr = pPr.find("w\\:numPr");
            if (numPr.length > 0) {
              const numId = numPr.find("w\\:numId").attr("w:val");
              const ilvl = numPr.find("w\\:ilvl").attr("w:val");

              const listData = extractListInfo(numId, ilvl, numberingMap);
              paragraphData.listData = listData;
            }
          }

          const nextChild = [];
          for (const run of $doc(child).find("w\\:r, w\\:drawing, w\\:pict")) {
            const runTag = $doc(run)[0].tagName;

            if (runTag === "w:r") {
              const runText = $doc(run).find("w\\:t").text();
              const rPr = $doc(run).find("w\\:rPr");
              const rStyleId = rPr.find("w\\:rStyle").attr("w:val");
              let runStyles = {};

              if (rStyleId && styleMap[rStyleId]) {
                runStyles = {
                  ...styleMap[rStyleId].runProperties,
                };
              }
              runStyles = {
                ...runStyles,
                ...extractRunStyles(rPr),
              };
              nextChild.push({
                text: runText,
                styles: runStyles,
              });
            } else if (runTag === "w:drawing") {
              const imageData = await parseDrawing(run);
              if (imageData) {
                children.push(imageData);
              }
            }
          }

          // Collect the text from the runs
          const paragraphText = nextChild
            .filter((child) => child.text)
            .map((child) => child.text)
            .join("");

          // Only add paragraphData if there is text content
          if (paragraphText.trim() !== "") {
            paragraphData.text = paragraphText;
            paragraphData.styles = {
              ...paragraphData.styles,
              ...nextChild.styles,
            };
            paragraphData.styleName = pStyleId;
            children.push(paragraphData);
          }
        } else if (tag === "w:tbl") {
          const tableData = {
            type: "table",
            rows: [],
          };

          for (const row of $doc(child).find("w\\:tr")) {
            const rowData = [];

            for (const cell of $doc(row).find("w\\:tc")) {
              const cellData = {
                type: "cell",
                content: await parseElement($doc(cell)),
              };
              rowData.push(cellData);
            }

            tableData.rows.push(rowData);
          }

          children.push(tableData);
        } else if (tag === "w:sectPr") {
          const sectionData = {
            type: "section",
            styles: {
              pageSize:
                $doc(child).find("w\\:pgSz").attr("w:w") +
                "x" +
                $doc(child).find("w\\:pgSz").attr("w:h"),
              margins: {
                top: $doc(child).find("w\\:pgMar").attr("w:top"),
                bottom: $doc(child).find("w\\:pgMar").attr("w:bottom"),
                left: $doc(child).find("w\\:pgMar").attr("w:left"),
                right: $doc(child).find("w\\:pgMar").attr("w:right"),
              },
            },
          };
          children.push(sectionData);
        }
      }

      return children;
    }

    function buildNumberingMap($numbering) {
      const abstractNumMap = {}; // Stores the abstract numbering definitions
      const numberingMap = {}; // Stores the final mapping of numId to its levels

      // Recursive function to resolve abstract numbering, including any references to other abstractNumIds
      function resolveAbstractNum(abstractNumId) {
        // If this abstractNumId is already resolved, return it
        if (abstractNumMap[abstractNumId]?.resolved) {
          return abstractNumMap[abstractNumId].levels;
        }

        const levels = abstractNumMap[abstractNumId]?.levels || [];

        // Check if this abstractNumId refers to another abstractNumId
        const referencedAbstractNumId =
          abstractNumMap[abstractNumId]?.referencedAbstractNumId;
        if (referencedAbstractNumId) {
          // Recursively resolve the referenced abstractNumId and merge the levels
          const referencedLevels = resolveAbstractNum(referencedAbstractNumId);
          Object.assign(levels, referencedLevels);
        }

        // Mark this abstractNumId as resolved to avoid circular references
        abstractNumMap[abstractNumId] = {
          levels,
          resolved: true,
        };

        return levels;
      }

      // Loop through each abstract numbering definition and build the abstractNumMap
      $numbering("w\\:abstractNum").each((_, abstractNum) => {
        const abstractNumId = $numbering(abstractNum).attr("w:abstractNumId");
        const levels = [];

        $numbering(abstractNum)
          .find("w\\:lvl")
          .each((_, level) => {
            const levelIndex = $numbering(level).attr("w:ilvl");
            const numFmt = $numbering(level).find("w\\:numFmt").attr("w:val");
            const lvlText = $numbering(level).find("w\\:lvlText").attr("w:val");

            // Get indentation spacing
            const indent = {};
            const indElement = $numbering(level).find("w\\:ind");
            if (indElement.length > 0) {
              indent.left = indElement.attr("w:left") || null;
              indent.hanging = indElement.attr("w:hanging") || null;
              indent.firstLine = indElement.attr("w:firstLine") || null;
            }

            levels[levelIndex] = {
              numFmt,
              lvlText,
              indent,
            };
          });

        // Check if this abstractNum refers to another abstractNum
        const referencedAbstractNumId = $numbering(abstractNum)
          .find("w\\:nsid")
          .attr("w:val");
        abstractNumMap[abstractNumId] = {
          levels,
          referencedAbstractNumId: referencedAbstractNumId || null,
          resolved: false, // Initially set as unresolved
        };
      });

      // Loop through each num element to build the numberingMap
      $numbering("w\\:num").each((_, num) => {
        const numId = $numbering(num).attr("w:numId");
        const abstractNumId = $numbering(num)
          .find("w\\:abstractNumId")
          .attr("w:val");

        if (abstractNumId) {
          numberingMap[numId] = resolveAbstractNum(abstractNumId);
        }
      });

      return numberingMap;
    }

    function removeNullFields(obj) {
      if (typeof obj === "object" && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (obj[key] === null || obj[key] === undefined) {
            delete obj[key];
          } else if (typeof obj[key] === "object") {
            removeNullFields(obj[key]); // Recursively clean nested objects
            // If the object becomes empty after cleaning, remove it
            if (Object.keys(obj[key]).length === 0) {
              delete obj[key];
            }
          }
        });
      }
    }

    async function parseDrawing(drawingElement) {
      const anchor = $doc(drawingElement).find("wp\\:anchor");
      const graphicData = $doc(drawingElement).find("a\\:graphicData");

      const imageData = {
        type: "image",
        src: null,
        position: {
          horizontal: null,
          vertical: null,
        },
        size: {
          width: null,
          height: null,
        },
        properties: {},
      };

      const positionH = anchor.find("wp\\:positionH wp\\:align").text();
      const positionV = anchor.find("wp\\:positionV wp\\:align").text();
      const cx = anchor.find("wp\\:extent").attr("cx");
      const cy = anchor.find("wp\\:extent").attr("cy");

      if (positionH && positionV) {
        imageData.position.horizontal = positionH;
        imageData.position.vertical = positionV;
      }

      if (cx && cy) {
        imageData.size.width = parseInt(cx, 10);
        imageData.size.height = parseInt(cy, 10);
      }

      const blip = graphicData.find("a\\:blip");
      const embedId = blip.attr("r:embed");

      if (embedId) {
        // Read the document relationships XML
        const relsXml = await zip
          .file("word/_rels/document.xml.rels")
          .async("text");
        const $rels = cheerio.load(relsXml, { xmlMode: true });

        // Find the image's target based on the embedId
        const target = $rels(`Relationship[Id="${embedId}"]`).attr("Target");
        if (target) {
          // Locate the image in the zip
          const imagePath = `word/${target}`;
          const imageFile = await zip.file(imagePath).async("blob"); // Get the image as a blob

          try {
            // Upload image to Cloudinary
            const uploadResult = await uploadImage(imageFile);
            imageData.src = uploadResult.url;
          } catch (error) {
            console.error("Error uploading image to Cloudinary:", error);
          }
        }
      }
      // Remove any null or empty fields from imageData
      removeNullFields(imageData);
      if (imageData.src) {
        return imageData;
      }
      return null;
    }

    function extractListInfo(numId, ilvl, numberingMap) {
      const levelInfo = numberingMap[numId] && numberingMap[numId][ilvl];

      return levelInfo
        ? {
            isBullet: levelInfo.numFmt === "bullet",
            bulletText: levelInfo.lvlText,
            level: ilvl,
            indent: levelInfo.indent,
          }
        : null;
    }

    const parsedContent = await parseElement($doc("w\\:document > w\\:body"));

    return mapSections(parsedContent);
  } catch (err) {
    console.error("Error extracting DOCX content:", err);
    throw err;
  }
}

function mapSections(paragraphs) {
  const sections = [];
  let currentSection = null;
  let currentSubsection = null;
  let preamble = { body: [] };
  let imageCounter = 1; // To count the figure numbers

  paragraphs.forEach((paragraph, index) => {
    const { styleName, text, type, ...otherProperties } = paragraph;
    // console.log(styleName, text);
    // Ignore paragraphs with no text
    if (type === "paragraph" && (!text || text.trim() === "")) {
      return;
    }

    // Handle Heading1 and TGTHEADING1 (Main Section)
    if (["Heading1", "TGTHEADING1", "a7", "1"].includes(styleName)) {
      // Push the current subsection to the current section if it has body content
      if (currentSubsection && currentSubsection.body.length > 0) {
        currentSection.body.push(currentSubsection);
      }

      // Push the current section to sections if it has body content
      if (currentSection && currentSection.body.length > 0) {
        sections.push(currentSection);
      }

      // Start a new section
      currentSection = {
        title: text,
        body: [],
      };

      // Reset current subsection
      currentSubsection = null;
    }
    // Handle TGTHEADING2 (Subsection)
    else if (
      ["TGTHEADING2", "Heading2", "21"].includes(styleName) &&
      currentSection
    ) {
      // Push the current subsection to the current section if it has body content
      if (currentSubsection && currentSubsection.body.length > 0) {
        currentSection.body.push(currentSubsection);
      }

      // Start a new subsection
      currentSubsection = {
        title: text,
        body: [],
      };
    }
    // Handle images and check for captions
    else if (type === "image") {
      const prevParagraph = paragraphs[index - 1];

      // If the previous paragraph is not a caption, insert a custom caption
      if (
        !prevParagraph ||
        (prevParagraph.styleName !== "Caption" &&
          prevParagraph.styleName !== "tgtcaption")
      ) {
        const customCaption = {
          type: "Text",
          value: `Figure ${imageCounter}: Custom caption for image`,
        };

        if (currentSubsection) {
          currentSubsection.body.push(customCaption);
        } else if (currentSection) {
          currentSection.body.push(customCaption);
        } else {
          preamble.body.push(customCaption);
        }

        imageCounter++; // Increment the image/figure counter
      }

      // Add the image to the appropriate section
      if (currentSubsection) {
        currentSubsection.body.push({
          value: text,
          type: "image",
          ...otherProperties,
        });
      } else if (currentSection) {
        currentSection.body.push({
          value: text,
          type: "image",
          ...otherProperties,
        });
      } else {
        preamble.body.push({
          value: text,
          type: "image",
          ...otherProperties,
        });
      }
    }
    // Handle adding paragraphs to subsections
    else if (currentSubsection) {
      currentSubsection.body.push({
        value: text,
        type: type === "paragraph" ? "Text" : type,
        ...otherProperties,
      });
    }
    // Handle adding paragraphs to sections
    else if (currentSection) {
      currentSection.body.push({
        value: text,
        type: type === "paragraph" ? "Text" : type,
        ...otherProperties,
      });
    }
    // Handle preamble before any sections
    else {
      preamble.body.push({
        value: text,
        type: type === "paragraph" ? "Text" : type,
        ...otherProperties,
      });
    }
  });

  // Push the last subsection to the current section if it has body content
  if (currentSubsection && currentSubsection.body.length > 0) {
    currentSection.body.push(currentSubsection);
  }

  // Push the current section to sections if it has body content
  if (currentSection && currentSection.body.length > 0) {
    sections.push(currentSection);
  }

  // Add preamble to sections if it has body content
  if (preamble.body.length > 0) {
    sections.unshift({
      title: "Cover Page",
      body: preamble.body,
    });
  }

  return sections;
}
