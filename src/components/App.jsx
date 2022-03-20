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
        saveAs(out, `${item[1]}.docx`);
      } catch (e) {
        console.log(e.message);
      }
    });
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
      console.log(list);
      list.forEach(generateDocument);
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
