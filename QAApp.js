/*
QAApp v1.0.0
GitHub: https://github.com/tanaikech/QAApp
*/

/**
 * Class object of OAApp.
 * 
 * OAApp v1.0.1
 * GitHub: https://github.com/tanaikech/OAApp
 */
class QAApp {

  /**
   * @constructor
   * 
   */
  constructor() {
    /** @private */
    this.apiKey;

    /** @private */
    this.accessToken = ScriptApp.getOAuthToken();

    /** @private */
    this.spreadsheetId;

    /** @private */
    this.dashboardSheet;

    /** @private */
    this.dataSheet;

    /** @private */
    this.databaseName;

    /** @private */
    this.corpusName;

    /** @private */
    this.documentName;

    /** @private */
    this.chunkNamePref;

    /** @private */
    this.dashboardSheetName = "dashboard"; // If you want to change the sheet name, please modify the sheet name and this value.

    /** @private */
    this.dataSheetName = "data"; // If you want to change the sheet name, please modify the sheet name and this value.

    /** @private */
    this.dataSheetId;

    /** @private */
    this.threshold;
  }

  /**
   * ### Description
   * Manage chunks. Set data on Spreadsheet to chunks in corpus.
   *
   * @returns {null}
   */
  manageChunks() {
    this.getSheets_();
    const obj = this.getValues_();
    this.createCorpus_();
    this.setChunks_(obj);
    return null;
  }

  /**
   * ### Description
   * Clear corpus. Delete all chunks in a corpus.
   *
   * @returns {null}
   */
  clearCorpus() {
    this.getSheets_();
    this.getValues_(false);

    const ui = SpreadsheetApp.getUi();
    const res = ui.alert("Clear corpus", `Do you want to clear all data in the corpus "${this.corpusName}"?`, ui.ButtonSet.YES_NO);
    if (res != ui.Button.YES) return;

    this.deleteCorpus_();
    this.createCorpus_();

    ui.alert("Done.");
    return null;
  }

  /**
   * ### Description
   * Get chunks from corpus and put them into Spreadsheet.
   *
   * @returns {null}
   */
  getChunksAndUpdateDataSheet() {
    this.getSheets_();
    this.getValues_(false);
    this.getChunks_();
    return null;
  }

  /**
   * ### Description
   * Generate answer to question.
   *
   * @param {Object} object Object including question.
   * @param {String} object.prompt Inputted question.
   * 
   * @returns {String} Return the generated answer.
   */
  generateAnswer(object) {
    this.getSheets_();
    this.getValues_(false);
    return this.searchChunks_(object);
  }

  /**
   * ### Description
   * Get sheets.
   * 
   * @return {void}
   * @private
   */
  getSheets_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.spreadsheetId = ss.getId();
    this.dashboardSheet = ss.getSheetByName(this.dashboardSheetName) || ss.insertSheet(this.dashboardSheetName);
    this.dataSheet = ss.getSheetByName(this.dataSheetName) || ss.insertSheet(this.dataSheetName);
    this.dataSheetId = this.dataSheet.getSheetId();
  }

  /**
   * ### Description
   * Get values from sheets.
   * 
   * @param {Boolean} getValues
   * 
   * @return {Object|null}
   * @private
   */
  getValues_(getValues = true) {
    const values = this.dashboardSheet.getRange("A2:C" + this.dashboardSheet.getLastRow()).getDisplayValues();
    const obj = values.reduce((o, [a, b]) => (o[a] = b, o), {});
    this.apiKey = obj.apiKey || "";
    this.databaseName = obj.databaseName;
    this.corpusName = obj.corpusName || `corpora/${this.databaseName}`;
    this.documentName = obj.documentName || `${this.corpusName}/documents/${this.databaseName}`;
    this.chunkNamePref = `${this.documentName}/chunks/`;
    this.threshold = obj.threshold ?? 0.7;

    const rule = SpreadsheetApp.newDataValidation().requireValueInList(["Update", "Delete"], true).setAllowInvalid(false).build();
    this.dataSheet.getRange("D2:D").setVerticalAlignment("middle").setHorizontalAlignment("center").setDataValidation(rule);

    if (getValues) {
      const dataRange = this.dataSheet.getRange("A2:E" + this.dataSheet.getLastRow());
      return { dataRange, data: dataRange.getDisplayValues() };
    }
    return null;
  }

  /**
   * ### Description
   * Create corpus and document in corpus.
   * 
   * @return {void}
   * @private
   */
  createCorpus_() {
    const CorporaApp = new CorporaApp_(this.accessToken);
    if (CorporaApp.getCorpora().some(e => e.name == this.corpusName)) {
      if (!CorporaApp.getDocuments(this.corpusName).some(e => e.name == this.documentName)) {
        CorporaApp.createDocument(this.corpusName, { name: this.documentName, displayName: this.databaseName });
        console.log(`Document "${this.documentName}" was created in the corpus "${this.corpusName}".`);
      }
    } else {
      CorporaApp.createCorpus({ name: this.corpusName, displayName: this.databaseName });
      console.log(`Corpus "${this.corpusName}" was created.`);
      CorporaApp.createDocument(this.corpusName, { name: this.documentName, displayName: this.databaseName });
      console.log(`Document "${this.documentName}" was created in the corpus "${this.corpusName}".`);
    }
  }

  /**
   * ### Description
   * Delete corpus.
   * 
   * @return {void}
   * @private
   */
  deleteCorpus_() {
    const CorporaApp = new CorporaApp_(this.accessToken);
    if (!CorporaApp.getCorpora().some(e => e.name == this.corpusName)) {
      console.warn(`Corpus of "${this.corpusName}" is not existing.`);
      return;
    }
    CorporaApp.deleteCorpus(this.corpusName, true);
    console.warn(`Corpus of "${this.corpusName}" was deleted.`);

    this.dataSheet.getRange("E2:E").clearContent();
  }

  /**
   * ### Description
   * Delete corpus.
   * 
   * @param {Object} obj
   * @param {String} obj.data Data from corpus.
   * 
   * @return {void}
   * @private
   */
  setChunks_(obj) {
    const { data } = obj;
    if (data.length == 0) {
      Browser.msgBox("No data.");
      return;
    }
    const { add, update, remove } = data.reduce((o, [date, q, a, status, id], i) => {
      if (q && a) {
        if (!id) {
          o.add.cells.push(`'${this.dataSheetName}'!E${i + 2}`);
          o.add.requests.push({
            parent: this.documentName,
            chunk: {
              data: { stringValue: `<Question>${q}</Question><Answer>${a}</Answer>` },
              customMetadata: [{ key: "date", stringValue: date }]
            }
          });
        } else if (status.toLowerCase() == "update") {
          o.update.cells.push(`D${i + 2}`);
          o.update.requests.push({
            chunk: {
              name: `${this.chunkNamePref}${id}`,
              data: { stringValue: `<Question>${q}</Question><Answer>${a}</Answer>` },
              customMetadata: [{ key: "date", stringValue: date }]
            },
            updateMask: "data,customMetadata"
          });
        } else if (status.toLowerCase() == "delete") {
          o.remove.cells.push({ deleteDimension: { range: { sheetId: this.dataSheetId, startIndex: i + 1, endIndex: i + 2, dimension: "ROWS" } } });
          o.remove.requests.push({ name: `${this.chunkNamePref}${id}` });
        }
      }
      return o;
    }, {
      add: { cells: [], requests: [] },
      update: { cells: [], requests: [] },
      remove: { cells: [], requests: [] }
    });

    const CorporaApp = new CorporaApp_(this.accessToken);
    if (add.requests.length > 0) {
      console.log(`${add.cells.length} items are added to "${this.documentName}" as chunks.`);
      try {
        const res = CorporaApp.setChunks(this.documentName, { requests: add.requests });
        const chunkNames = res.flatMap(e => JSON.parse(e.getContentText()).chunks.map(({ name }) => name.split("/").pop()));
        Sheets.Spreadsheets.Values.batchUpdate({ data: chunkNames.map((id, i) => ({ range: add.cells[i], values: [[id]] })), valueInputOption: "USER_ENTERED" }, this.spreadsheetId);
      } catch ({ stack }) {
        console.error(stack);
        if (stack.includes("At least one text exceeds the token limit by")) {
          Browser.msgBox(`The text size of Q&A of columns "B" and "C" is large. Please reduce the text size and try again.`);
        } else {
          Browser.msgBox(stack);
        }
        return;
      }
    }
    if (update.requests.length > 0) {
      console.log(`${update.cells.length} chunks in "${this.documentName}" are updated.`);
      try {
        CorporaApp.updateChunks(this.documentName, { requests: update.requests });
        this.dataSheet.getRangeList(update.cells).setValue("");
      } catch ({ stack }) {
        console.error(stack);
        if (stack.includes("At least one text exceeds the token limit by")) {
          Browser.msgBox(`The text size of Q&A of columns "B" and "C" is large. Please reduce the text size and try again.`);
        } else {
          Browser.msgBox(stack);
        }
        return;
      }
    }
    if (remove.requests.length > 0) {
      console.log(`${remove.cells.length} chunks in "${this.documentName}" are deleted.`);
      try {
        CorporaApp.deleteChunks(this.documentName, { requests: remove.requests });
        this.dataSheet.insertRowsAfter(this.dataSheet.getMaxRows(), remove.cells.length);
        Sheets.Spreadsheets.batchUpdate({ requests: remove.cells.reverse() }, this.spreadsheetId);
      } catch ({ message, stack }) {
        console.error(stack);
        const o = JSON.parse(message);
        if (o.error.status == "NOT_FOUND") {
          Browser.msgBox("Chunk IDs for deleting were not found.");
        } else {
          Browser.msgBox(stack);
        }
        return;
      }
    }
    Browser.msgBox("Done.");
  }

  /**
   * ### Description
   * Get chunks from corpus.
   * 
   * @return {void}
   * @private
   */
  getChunks_() {
    const CorporaApp = new CorporaApp_(this.accessToken);
    const ar = CorporaApp.getChunks(this.documentName);
    console.log(`${ar.length} chunks are existing in "${this.documentName}".`);
    if (ar.length == 0) return;
    const values = ar.map(({ data: { stringValue }, customMetadata, name }) => {
      const date = customMetadata.find(({ key }) => key == "date");
      let q = "";
      let a = "";
      if (stringValue) {
        q = stringValue.match(/<Question>(.*)<\/Question>/)[1].trim();
        a = stringValue.match(/<Answer>(.*)<\/Answer>/)[1].trim();
      }
      return [date ? (date.stringValue || "") : "", q, a, "", name.split("/").pop()];
    });
    if (values.length > 0) {
      this.dataSheet.getRange("A2:E" + this.dataSheet.getLastRow()).clearContent();
      this.dataSheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    } else {
      console.warn(`No values are existing in "${this.documentName}".`);
    }
  }

  /**
   * ### Description
   * Get chunks from corpus.
   * 
   * @param {Object} object Object including question.
   * @param {String} object.prompt Inputted question.
   * 
   * @return {String} Return generated answer.
   * @private
   */
  searchChunks_(object) {
    const { prompt } = object;
    if (prompt == "") {
      Browser.msgBox(`Please put a question.`);
      return;
    }
    const threshold = 0.7;
    const CorporaApp = new CorporaApp_(this.accessToken);
    const res = CorporaApp.searchQueryFromDocument(this.documentName, { query: prompt, resultsCount: 100 });
    const text = res.getContentText();
    if (res.getResponseCode() != 200) {
      Browser.msgBox(text);
      console.error(text);
      return;
    }
    const ar = JSON.parse(text);

    if (!this.apiKey) {
      const msg = "Please set your API key for using Gemini API.";
      Browser.msgBox(msg);
      console.error(msg);
      return;
    }

    const g = new GeminiWithFiles({ apiKey: this.apiKey });
    let q = "";
    if (ar.relevantChunks && ar.relevantChunks.length > 0) {
      console.log(`${ar.relevantChunks.length} chunks were retrieved from Corpus with the threshold ${this.threshold}. These will be used with generate content with Gemini API.`);
      const temp = ar.relevantChunks.filter(({ chunkRelevanceScore }) => chunkRelevanceScore > threshold);
      const data = temp.map(({ chunk: { data: { stringValue }, createTime, updateTime } }) => {
        const q = stringValue.match(/<Question>(.*)<\/Question>/)[1].trim();
        const a = stringValue.match(/<Answer>(.*)<\/Answer>/)[1].trim();
        return { q, a, createTime, updateTime };
      });
      const jsonSchema = {
        description: "This data includes past experiences, and this can be used to help answer the main question. Generate an answer using this data.",
        type: "array",
        items: {
          type: "object",
          properties: {
            q: { description: "Question related to the main question.", type: "string" },
            a: { description: "Answer to question.", type: "string" },
            createTime: { description: "Created time of question and answer.", type: "string" },
            updateTime: { description: "Updated time of question and answer.", type: "string" },
          },
        },
      };
      q = [
        `Run the following step.`,
        `1. Understand the main question. The main question is as follows.`,
        `<MainQuestion>${prompt}</MainQuestion>`,
        `2. Understand the reference data to help understand the main question. JSON schema of the reference data of the "data.txt" file is as follows. This data includes past experiences, and this can be used to help answer the main question. Basically, generate an answer using this data.`,
        `<JSONSchema>${JSON.stringify(jsonSchema)}</JSONSchema>`,
        `3. Concisely return the answer to the main question as a text based on the uploaded data of the "data.txt" file and your knowledge by considering "IMPORTANT".`,
        `<IMPORTANT>`,
        `If a clear answer cannot be generated from the reference data from the "data.txt" file, generate an answer based on only your knowledge by ignoring the reference data from the "data.txt" file.`,
        `If you need more information to generate a clear answer, return only the response "Cannot answer by lack of information.".`,
        `If you have no clear answer to the main question even when both the reference data from the "data.txt" file and your knowledge are used, return the response "Cannot answer by no information.".`,
        `The generated content must not include JSON data.`,
        `</IMPORTANT>`,
      ].join("\n");
      const blob = Utilities.newBlob(JSON.stringify(data), MimeType.PLAIN_TEXT, "data.txt");
      const fileList = g.setBlobs([blob]).uploadFiles();
      g.withUploadedFilesByGenerateContent(fileList);
    } else {
      q = [
        `Run the following step.`,
        `1. Understand the main question. The main question is as follows.`,
        `<MainQuestion>${prompt}</MainQuestion>`,
        `2. Concisely return the answer to the main question as a text based on your knowledge.`,
      ].join("\n");
    }
    return g.generateContent({ q });
  }

}
