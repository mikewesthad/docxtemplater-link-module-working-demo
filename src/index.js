import DocXTemplater from "docxtemplater";
import JSZipUtils from "jszip-utils";
import JSZip from "jszip";
import saveAs from "file-saver";
import LinkModule from "docxtemplater-link-module";
import templateUrl from "./simple-href-template.docx";

const linkModule = new LinkModule();
const loadFile = (url, callback) => JSZipUtils.getBinaryContent(url, callback);
const downloadButton = document.querySelector("#download");

downloadButton.addEventListener("click", () => {
  loadFile(templateUrl, function(error, content) {
    if (error) {
      alert(e.message);
      throw error;
    }

    const zip = new JSZip(content);
    const doc = new DocXTemplater()
      .attachModule(linkModule)
      .load(zip)
      .setData({
        link: {
          text: "dolor sit",
          url: "http://google.com"
        }
      });

    try {
      doc.render();
    } catch (error) {
      const e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties
      };
      console.log(JSON.stringify({ error: e }));
      alert(e.message);
      throw error;
    }

    const out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    saveAs(out, "output.docx");
  });
});
