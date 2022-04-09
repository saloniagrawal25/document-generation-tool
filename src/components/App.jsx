import React from "react";
import { read, utils } from "xlsx";
import DocxTemplater from "docxtemplater";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

const App = () => {
  const finalZip = new PizZip();

  const generateDocument = (item) => {
    loadFile("http://127.0.0.1:8080/Template.docx", function (error, content) {
      try {
        if (error) {
          throw error;
        }
        const zip = new PizZip(content);
        const doc = new DocxTemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        doc.render(item);
        const out = doc.getZip().generate({
          type: "blob",
          mimeType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        blobToBase64(out, function (binaryData) {
          finalZip.file(`${item[1]}.docx`, binaryData, { base64: true });
        });
        //saveAs(out, `${item[1]}.docx`);
      } catch (e) {
        console.log(e.message);
      }
    });
  };

  const blobToBase64 = (blob, callback) => {
    var reader = new FileReader();
    reader.onload = function () {
      var dataUrl = reader.result;
      var base64 = dataUrl.split(",")[1];
      callback(base64);
    };
    reader.readAsDataURL(blob);
  };

  const onButtonClick = (e) => {
    const [file] = e.target.files;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = read(bstr, { type: "binary" });
      const wsName = wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      const list = utils.sheet_to_json(ws, { header: 1 });
      list.forEach(generateDocument);
      PizZipUtils.getBinaryContent();
      const content = finalZip.generate({
        type: "blob",
      });
      saveAs(content, "output.zip");
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div>
      <input type="file" id="file" onChange={onButtonClick} />
    </div>
  );
};
export default App;
