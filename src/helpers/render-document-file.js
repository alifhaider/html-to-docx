/* eslint-disable no-await-in-loop */
/* eslint-disable no-case-declarations */
import { fragment } from 'xmlbuilder2';
import VNode from 'virtual-dom/vnode/vnode';
import VText from 'virtual-dom/vnode/vtext';
import isVNode from 'virtual-dom/vnode/is-vnode';
import isVText from 'virtual-dom/vnode/is-vtext';
// eslint-disable-next-line import/no-named-default
import { default as HTMLToVDOM } from 'html-to-vdom';
import sizeOf from 'image-size';
import imageToBase64 from 'image-to-base64';

// FIXME: remove the cyclic dependency
import * as xmlBuilder from './xml-builder';
import namespaces from '../namespaces';
import { imageType, internalRelationship } from '../constants';
import { vNodeHasChildren } from '../utils/vnode';
import { isValidUrl } from '../utils/url';
import { getMimeType } from '../utils/image';

const convertHTML = HTMLToVDOM({
  VNode,
  VText,
});

// eslint-disable-next-line consistent-return, no-shadow
export const buildImage = async (docxDocumentInstance, vNode, maximumWidth = null) => {
  let response = null;
  let base64Uri = null;
  try {
    const imageSource = vNode.properties.src;
    if (isValidUrl(imageSource)) {
      const base64String = await imageToBase64(imageSource).catch((error) => {
        // eslint-disable-next-line no-console
        console.warn(`skipping image download and conversion due to ${error}`);
        return null;
      });

      if (base64String) {
        const mimeType = getMimeType(imageSource, base64String);
        base64Uri = `data:${mimeType};base64, ${base64String}`;
      } else {
        console.error(`[ERROR] buildImage: Failed to convert URL to base64`);
      }
    } else {
      base64Uri = decodeURIComponent(vNode.properties.src);
    }

    if (base64Uri) {
      response = docxDocumentInstance.createMediaFile(base64Uri);
    } else {
      console.error(`[ERROR] buildImage: No valid base64Uri generated`);
      return null;
    }
  } catch (error) {
    console.error(`[ERROR] buildImage: Error during image processing:`, error);
    return null;
  }

  if (response) {
    try {
      docxDocumentInstance.zip
        .folder('word')
        .folder('media')
        .file(response.fileNameWithExtension, Buffer.from(response.fileContent, 'base64'), {
          createFolders: false,
        });

      const documentRelsId = docxDocumentInstance.createDocumentRelationships(
        docxDocumentInstance.relationshipFilename,
        imageType,
        `media/${response.fileNameWithExtension}`,
        internalRelationship
      );

      const imageBuffer = Buffer.from(response.fileContent, 'base64');
      const imageProperties = sizeOf(imageBuffer);

      const imageFragment = await xmlBuilder.buildParagraph(
        vNode,
        {
          type: 'picture',
          inlineOrAnchored: true,
          relationshipId: documentRelsId,
          ...response,
          description: vNode.properties.alt,
          maximumWidth: maximumWidth || docxDocumentInstance.availableDocumentSpace,
          originalWidth: imageProperties.width,
          originalHeight: imageProperties.height,
        },
        docxDocumentInstance
      );

      return imageFragment;
    } catch (error) {
      console.error(`[ERROR] buildImage: Error during XML generation:`, error);
      return null;
    }
  } else {
    console.error(`[ERROR] buildImage: No response from createMediaFile`);
    return null;
  }
};

export const buildList = async (vNode, docxDocumentInstance, xmlFragment) => {
  // Helper to merge parent styles/attributes into child node
  function mergeListParentProps(parentProps, currentVNode) {
    return {
      attributes: {
        ...(parentProps?.attributes || {}),
        ...(currentVNode?.properties?.attributes || {}),
      },
      style: {
        ...(parentProps?.style || {}),
        ...(currentVNode?.properties?.style || {}),
      },
    };
  }

  let vNodeObjects = [
    {
      node: vNode,
      level: 0,
      type: vNode.tagName,
      numberingId: docxDocumentInstance.createNumbering(vNode.tagName, vNode.properties),
      parentProps: vNode.properties,
    },
  ];

  while (vNodeObjects.length) {
    const tempVNodeObject = vNodeObjects.shift();
    const currentVNode = tempVNodeObject.node;
    const { parentProps } = tempVNodeObject;

    // text nodes will be handled in their parent
    if (!isVText(currentVNode)) {
      if (currentVNode.tagName === 'li') {
        const mergedProps = mergeListParentProps(parentProps, currentVNode);

        // Paragraph for the list item and its children (i.e <span>, <strong>)
        const paragraphVNode = new VNode('p', { ...mergedProps }, currentVNode.children);

        const paragraphFragment = await xmlBuilder.buildParagraph(
          paragraphVNode,
          {
            numbering: {
              levelId: tempVNodeObject.level,
              numberingId: tempVNodeObject.numberingId,
            },
          },
          docxDocumentInstance
        );
        xmlFragment.import(paragraphFragment);
      }

      if (currentVNode.children && currentVNode.children.length) {
        const newVNodeObjects = currentVNode.children
          .filter((childVNode) => !isVText(childVNode))
          .map((childVNode) => {
            if (['ul', 'ol'].includes(childVNode.tagName)) {
              // Nested list
              return {
                node: childVNode,
                level: tempVNodeObject.level + 1,
                type: childVNode.tagName,
                numberingId: docxDocumentInstance.createNumbering(
                  childVNode.tagName,
                  childVNode.properties
                ),
                parentProps: mergeListParentProps(parentProps, childVNode),
              };
            }
            if (childVNode.tagName === 'li') {
              return {
                node: childVNode,
                level: tempVNodeObject.level,
                type: tempVNodeObject.type,
                numberingId: tempVNodeObject.numberingId,
                parentProps,
              };
            }
            return null;
          })
          .filter(Boolean);

        if (newVNodeObjects.length > 0) {
          vNodeObjects = newVNodeObjects.concat(vNodeObjects);
        }
      }
    }
  }
};

async function findXMLEquivalent(docxDocumentInstance, vNode, xmlFragment) {
  if (
    vNode.tagName === 'div' &&
    (vNode.properties.attributes.class === 'page-break' ||
      (vNode.properties.style && vNode.properties.style['page-break-after']))
  ) {
    const paragraphFragment = fragment({ namespaceAlias: { w: namespaces.w } })
      .ele('@w', 'p')
      .ele('@w', 'r')
      .ele('@w', 'br')
      .att('@w', 'type', 'page')
      .up()
      .up()
      .up();

    xmlFragment.import(paragraphFragment);
    return;
  }

  switch (vNode.tagName) {
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6':
      const headingFragment = await xmlBuilder.buildParagraph(
        vNode,
        {
          paragraphStyle: `Heading${vNode.tagName[1]}`,
        },
        docxDocumentInstance
      );
      xmlFragment.import(headingFragment);
      return;
    case 'span':
    case 'strong':
    case 'b':
    case 'em':
    case 'i':
    case 'u':
    case 'ins':
    case 'strike':
    case 'del':
    case 's':
    case 'sub':
    case 'sup':
    case 'mark':
    case 'p':
    case 'a':
    case 'blockquote':
    case 'code':
    case 'pre':
      const paragraphFragment = await xmlBuilder.buildParagraph(vNode, {}, docxDocumentInstance);
      xmlFragment.import(paragraphFragment);
      return;
    case 'figure':
      if (vNodeHasChildren(vNode)) {
        // eslint-disable-next-line no-plusplus
        for (let index = 0; index < vNode.children.length; index++) {
          const childVNode = vNode.children[index];
          if (childVNode.tagName === 'table') {
            const tableFragment = await xmlBuilder.buildTable(
              childVNode,
              {
                maximumWidth: docxDocumentInstance.availableDocumentSpace,
                rowCantSplit: docxDocumentInstance.tableRowCantSplit,
              },
              docxDocumentInstance
            );
            xmlFragment.import(tableFragment);
            // Adding empty paragraph for space after table only if the option is enabled
            if (docxDocumentInstance.addSpacingAfterTable) {
              const emptyParagraphFragment = await xmlBuilder.buildParagraph(null, {});
              xmlFragment.import(emptyParagraphFragment);
            }
          } else if (childVNode.tagName === 'img') {
            const imageFragment = await buildImage(docxDocumentInstance, childVNode);
            if (imageFragment) {
              // Add lineRule attribute for consistency
              // Direct image processing includes this attribute, but HTML image processing was missing it
              // This ensures both processing paths generate identical XML structure
              imageFragment
                .first()
                .first()
                .att(
                  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                  'lineRule',
                  'auto'
                );
              xmlFragment.import(imageFragment);
            } else {
              console.log(
                `[DEBUG] findXMLEquivalent: buildImage returned null/undefined in figure`
              );
            }
          }
        }
      }
      return;
    case 'table':
      const tableFragment = await xmlBuilder.buildTable(
        vNode,
        {
          maximumWidth: docxDocumentInstance.availableDocumentSpace,
          rowCantSplit: docxDocumentInstance.tableRowCantSplit,
        },
        docxDocumentInstance
      );
      xmlFragment.import(tableFragment);
      // Adding empty paragraph for space after table only if the option is enabled
      if (docxDocumentInstance.addSpacingAfterTable) {
        const emptyParagraphFragment = await xmlBuilder.buildParagraph(null, {});
        xmlFragment.import(emptyParagraphFragment);
      }
      return;
    case 'ol':
    case 'ul':
      await buildList(vNode, docxDocumentInstance, xmlFragment);
      return;
    case 'img':
      const imageFragment = await buildImage(docxDocumentInstance, vNode);
      if (imageFragment) {
        // Add lineRule attribute for consistency
        // Direct image processing includes this attribute, but HTML image processing was missing it
        // This ensures both processing paths generate identical XML structure
        imageFragment
          .first()
          .first()
          .att('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'lineRule', 'auto');
        xmlFragment.import(imageFragment);
      } else {
        console.log(`[DEBUG] findXMLEquivalent: buildImage returned null/undefined`);
      }
      return;
    case 'br':
      const linebreakFragment = await xmlBuilder.buildParagraph(null, {});
      xmlFragment.import(linebreakFragment);
      return;
    case 'head':
      return;
  }
  if (vNodeHasChildren(vNode)) {
    // eslint-disable-next-line no-plusplus
    for (let index = 0; index < vNode.children.length; index++) {
      const childVNode = vNode.children[index];
      // eslint-disable-next-line no-use-before-define
      await convertVTreeToXML(docxDocumentInstance, childVNode, xmlFragment);
    }
  }
}

// eslint-disable-next-line consistent-return
export async function convertVTreeToXML(docxDocumentInstance, vTree, xmlFragment) {
  if (!vTree) {
    // eslint-disable-next-line no-useless-return
    return '';
  }
  if (Array.isArray(vTree) && vTree.length) {
    // eslint-disable-next-line no-plusplus
    for (let index = 0; index < vTree.length; index++) {
      const vNode = vTree[index];
      await convertVTreeToXML(docxDocumentInstance, vNode, xmlFragment);
    }
  } else if (isVNode(vTree)) {
    await findXMLEquivalent(docxDocumentInstance, vTree, xmlFragment);
  } else if (isVText(vTree)) {
    const paragraphFragment = await xmlBuilder.buildParagraph(vTree, {}, docxDocumentInstance);
    xmlFragment.import(paragraphFragment);
  }
  return xmlFragment;
}

async function renderDocumentFile(docxDocumentInstance) {
  const vTree = convertHTML(docxDocumentInstance.htmlString);

  const xmlFragment = fragment({ namespaceAlias: { w: namespaces.w } });

  const populatedXmlFragment = await convertVTreeToXML(docxDocumentInstance, vTree, xmlFragment);

  return populatedXmlFragment;
}

export default renderDocumentFile;
