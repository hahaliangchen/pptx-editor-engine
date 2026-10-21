import PptViewer from "../../src/index";
import { createBrowserDesignConverter, jsonToHtml } from "../../src/design/browser";
import { DesignBridge } from "../../src/page/design-bridge";
import { measureHtmlInBrowser } from "../../src/design/measure-browser";
import { base64ToBytes, bytesToBase64 } from "../../src/design/bytes";
import { nodeText } from "../../src/design/model";

Object.assign(window, { designTest: { PptViewer, createBrowserDesignConverter, jsonToHtml, DesignBridge,
  measureHtmlInBrowser, base64ToBytes, bytesToBase64, nodeText } });
