/**
 * @license
 * Copyright 2025 Abusanan Chemi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

const SPREADSHEET_ID = '10GI4SN_i-dmzMGmIfd5abhQXc7ujgLt5U8jzANH-pCw';
const SHEET_NAME = 'info';

/**
 * Serves the HTML file for the web app.
 * @returns {HtmlOutput} The HTML output to be rendered.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("ARABIC Story eBook")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper function to open the spreadsheet by ID.
 * @returns {Spreadsheet} The Google Spreadsheet object.
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Gets or creates the 'info' sheet. Initialises it with headers and default books if it is empty.
 * @returns {Sheet} The sheet object.
 */
function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // If the sheet has no rows (is empty), initialize it with headers and default books
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'id', 'title', 'author', 'cover', 'coverType', 
      'description', 'tags', 'rating', 'popularity', 
      'releaseDate', 'canvaLink'
    ]);
    
    const defaultBooks = getDefaultBooks();
    defaultBooks.forEach(book => {
      sheet.appendRow([
        book.id,
        book.title,
        book.author,
        book.cover,
        book.coverType,
        book.description,
        book.tags.join(', '),
        book.rating,
        book.popularity,
        book.releaseDate,
        book.canvaLink
      ]);
    });
  }
  
  return sheet;
}

/**
 * Reads all book rows from Google Sheet.
 * @returns {Array<Object>} List of books.
 */
function getBooks() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const books = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const book = {};
      
      headers.forEach((header, index) => {
        if (!header || header.toString().trim() === '') return;
        
        let val = row[index];
        const headerStr = header.toString().trim();
        
        if (headerStr === 'id' || headerStr === 'rating' || headerStr === 'popularity') {
          val = val ? Number(val) : 0;
        } else if (headerStr === 'tags') {
          val = val ? val.toString().split(',').map(t => t.trim()).filter(t => t !== '') : [];
        } else if (headerStr === 'releaseDate') {
          if (val instanceof Date) {
            try {
              val = val.toISOString().split('T')[0];
            } catch (e) {
              val = val.toString();
            }
          } else {
            val = val ? val.toString() : '';
          }
        } else {
          val = val ? val.toString() : '';
        }
        
        book[headerStr] = val;
      });
      
      // Storybook ให้แสดงของที่แนบลิงก์เท่านั้น
      if (book.canvaLink && book.canvaLink.toString().trim() !== '') {
        books.push(book);
      }
    }
    return books;
  } catch (error) {
    Logger.log("Error in getBooks: " + error.toString());
    throw new Error(error.toString());
  }
}

/**
 * Adds a new book to the Google Sheet and returns the updated book list.
 * @param {Object} book The book details.
 * @returns {Array<Object>} Updated list of books.
 */
function addBook(book) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let nextId = 1;
    if (data.length > 1) {
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      if (idIndex !== -1) {
        const ids = data.slice(1).map(row => Number(row[idIndex])).filter(val => !isNaN(val));
        if (ids.length > 0) {
          nextId = Math.max(...ids) + 1;
        }
      }
    }
    sheet.appendRow([
      nextId,
      book.title,
      book.author,
      book.cover || book.canvaLink,
      book.coverType || 'iframe',
      book.description || '',
      book.tags || '',
      Number(book.rating) || 5,
      Number(book.popularity) || 90,
      book.releaseDate || new Date().toISOString().split('T')[0],
      book.canvaLink,
      new Date().getTime() // L (Timestamp)
    ]);
    return getBooks();
  } catch (error) {
    Logger.log("Error in addBook: " + error.toString());
    throw new Error(error.toString());
  }
}

/**
 * Deletes a book row by its ID and returns the updated list.
 * @param {number} id The book ID to delete.
 * @returns {Array<Object>} Updated list of books.
 */
function deleteBook(id) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return getBooks();
  } catch (error) {
    Logger.log("Error in deleteBook: " + error.toString());
    throw new Error(error.toString());
  }
}

/**
 * Updates an existing book row in Google Sheet.
 * @param {Object} updatedBook The book details to update.
 * @returns {Array<Object>} Updated list of books.
 */
function updateBook(updatedBook) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error("Sheet is empty");
    
    const headers = data[0];
    const idIndex = headers.indexOf('id');
    if (idIndex === -1) throw new Error("Header 'id' not found");
    
    let rowNum = -1;
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][idIndex]) === Number(updatedBook.id)) {
        rowNum = i + 1; // 1-based index (header is row 1, data starts at row 2)
        break;
      }
    }
    
    if (rowNum === -1) throw new Error("Book not found with ID: " + updatedBook.id);
    
    // Create new row data based on sheet headers
    const newRow = headers.map((header, colIndex) => {
      if (header === 'id') return Number(updatedBook.id);
      if (header === 'rating') return Number(updatedBook.rating);
      if (header === 'popularity') return Number(updatedBook.popularity);
      if (header === 'tags') return Array.isArray(updatedBook.tags) ? updatedBook.tags.join(', ') : updatedBook.tags;
      
      // Preserve existing timestamp on update if not provided in updatedBook (Column L is index 11)
      if ((header === 'timestamp' || colIndex === 11) && updatedBook[header] === undefined) {
        return data[rowNum - 1][colIndex] || '';
      }
      
      return updatedBook[header] !== undefined ? updatedBook[header] : '';
    });
    
    sheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
    return getBooks();
  } catch (error) {
    Logger.log("Error in updateBook: " + error.toString());
    throw new Error(error.toString());
  }
}

/**
 * Initial fallback book list data.
 * @returns {Array<Object>} Default books data.
 */
function getDefaultBooks() {
  const defaults = [
    { id: 1, title: 'บทเรียนชีวิตจากอีดิลอัฎฮา การอดทนและความหวัง', author: 'Halwanee solaeh', cover: 'https://www.canva.com/design/DAGup-brpYQ/rW7Af-p-5Id56934RVG8Hw/view?embed', coverType: 'iframe', description: 'ค้นพบบทเรียนอันลึกซึ้งเกี่ยวกับการอดทนและความหวัง ผ่านเรื่องราวของวันอีดิลอัฎฮาอันประเสริฐ', tags: ['ศาสนา', 'บทเรียน'], rating: 5, popularity: 95, releaseDate: '2025-08-20', canvaLink: 'https://www.canva.com/design/DAGup-brpYQ/rW7Af-p-5Id56934RVG8Hw/view?embed' },
    { id: 2, title: 'อีดิลอัฎฮา วันแห่งความเมตตาและการแบ่งปัน', author: 'Almalina sama-ae', cover: 'https://www.canva.com/design/DAGxnUkJksI/NHPBHQ-zz-vetsmPKWtB-g/view?embed', coverType: 'iframe', description: 'เฉลิมฉลองวันแห่งความเมตตาและการแบ่งปัน เรียนรู้ความสำคัญของการให้ in วันอีดิลอัฎฮา', tags: ['ครอบครัว', 'แบ่งปัน'], rating: 4, popularity: 88, releaseDate: '2025-08-15', canvaLink: 'https://www.canva.com/design/DAGxnUkJksI/NHPBHQ-zz-vetsmPKWtB-g/view?embed' },
    { id: 3, title: 'อีดิลอัฎฮา อันประเสริฐ คุณค่าศรัทธาและความยำเกรง', author: 'Anees samae', cover: 'https://www.canva.com/design/DAGxnXoz9gA/kwe68FwrFyH57baA8iJl-Q/view?embed', coverType: 'iframe', description: 'สำรวจคุณค่าแห่งศรัทธาและความยำเกรงต่อพระผู้เป็นเจ้าในเทศกาลอีดิลอัฎฮาที่เปี่ยมด้วยความหมาย', tags: ['ศรัทธา', 'จิตวิญญาณ'], rating: 5, popularity: 92, releaseDate: '2025-08-10', canvaLink: 'https://www.canva.com/design/DAGxnXoz9gA/kwe68FwrFyH57baA8iJl-Q/view?embed' },
    { id: 4, title: 'อีดิลอัฎฮาบทเรียนแห่งการเสียสละอันยิ่งใหญ่', author: 'Farial jumi', cover: 'https://www.canva.com/design/DAGvvN9_Kd8/GWmJb5uh3LoyQMUigIWOdg/view?embed', coverType: 'iframe', description: 'เรียนรู้การเสียสละอันยิ่งใหญ่ผ่านเรื่องราวของนบีอิบรอฮีมและนบีอิสมาอีลในวันอีดิลอัฎฮา', tags: ['เสียสละ', 'ประวัติศาสตร์'], rating: 5, popularity: 98, releaseDate: '2025-09-01', canvaLink: 'https://www.canva.com/design/DAGvvN9_Kd8/GWmJb5uh3LoyQMUigIWOdg/view?embed' },
    { id: 5, title: 'กุรบาน สัญลักษณ์แห่งความภักดีต่อพระผู้เป็นเจ้า', author: 'Bayan hayimaming', cover: 'https://www.canva.com/design/DAGupwxWZ3o/6YmOY1JeS2E3Wls_oGCO7Q/view?embed', coverType: 'iframe', description: 'เข้าใจความหมายที่แท้จริงของ "กุรบาน" สัญลักษณ์แห่งความรักและความภักดีอันสูงสุด', tags: ['กุรบาน', 'ภักดี'], rating: 4, popularity: 85, releaseDate: '2025-07-28', canvaLink: 'https://www.canva.com/design/DAGupwxWZ3o/6YmOY1JeS2E3Wls_oGCO7Q/view?embed' },
    { id: 6, title: 'รำลึกนบีอิบรอฮีมและนบีอิสมาอีลแห่งความเชื่อมั่น', author: 'Nur-eman maha', cover: 'https://www.canva.com/design/DAGupy17soc/-DJUzBeTcZNE9UyDAhOZvA/view?embed', coverType: 'iframe', description: 'ย้อนรำลึกถึงเรื่องราวแห่งความเชื่อมั่นของสองนบีผู้ยิ่งใหญ่ ต้นแบบแห่งการศรัทธา', tags: ['ประวัติศาสตร์', 'ศรัทธา'], rating: 5, popularity: 93, releaseDate: '2025-08-25', canvaLink: 'https://www.canva.com/design/DAGupy17soc/-DJUzBeTcZNE9UyDAhOZvA/view?embed' },
    { id: 7, title: 'ฮัจญ์และอีดิลอัฎฮา เส้นทางสู่การชำระหัวใจ', author: 'Wannasneen Waedueramae', cover: 'https://www.canva.com/design/DAGup7V72H0/pYHrwILMeM_VdBzaX6aOvg/view?embed', coverType: 'iframe', description: 'การเดินทางสู่พิธีฮัจญ์และวันอีด คือเส้นทางแห่งการขัดเกลาและชำระล้างจิตใจให้บริสุทธิ์', tags: ['ฮัจญ์', 'จิตวิญญาณ'], rating: 5, popularity: 90, releaseDate: '2025-08-05', canvaLink: 'https://www.canva.com/design/DAGup7V72H0/pYHrwILMeM_VdBzaX6aOvg/view?embed' },
    { id: 8, title: 'อีดิลอัฎฮา แรงบันดาลใจในการเปลี่ยนแปลงสู่ความดี', author: 'Afreen Maeroh', cover: 'https://www.canva.com/design/DAGsrpvnvoM/pu2vniGrDsf5JRLMn3lWng/view?embed', coverType: 'iframe', description: 'รับแรงบันดาลใจจากวันอีดเพื่อเริ่มต้นการเปลี่ยนแปลงตนเองไปสู่สิ่งที่ดีกว่าในทุกๆ วัน', tags: ['แรงบันดาลใจ', 'พัฒนาตนเอง'], rating: 4, popularity: 89, releaseDate: '2025-08-18', canvaLink: 'https://www.canva.com/design/DAGsrpvnvoM/pu2vniGrDsf5JRLMn3lWng/view?embed' },
    { id: 9, title: 'อีดิลอัฎฮา สายใยแห่งครอบครัวและสังคมมุสลิม', author: 'Nur-imee  Chedo', cover: 'https://www.canva.com/design/DAGt90c6HlU/uGhz2aNtluB1ZFalzeQqmA/watch?embed', coverType: 'iframe', description: 'เฉลิมฉลองและกระชับสายสัมพันธ์ในครอบครัวและสังคมมุสลิมให้แน่นแฟ้นยิ่งขึ้นในวันอีด', tags: ['ครอบครัว', 'สังคม'], rating: 5, popularity: 96, releaseDate: '2025-08-28', canvaLink: 'https://www.canva.com/design/DAGt90c6HlU/uGhz2aNtluB1ZFalzeQqmA/watch?embed' },
    { id: 10, title: 'อีดิลอัฎฮา วันแห่งการรวมใจในละหมาด', author: 'Matsanee    Kodcharit', cover: 'https://www.canva.com/design/DAGt9x9Z3q4/RbQUqQnXhsYRHmmOT9GYiQ/view?embed', coverType: 'iframe', description: 'สัมผัสพลังแห่งการรวมใจเป็นหนึ่งเดียวในการละหมาดและการรำลึกถึงอัลลอฮ์ในวันอันประเสริฐ', tags: ['ละหมาด', 'สังคม'], rating: 4, popularity: 87, releaseDate: '2025-07-20', canvaLink: 'https://www.canva.com/design/DAGt9x9Z3q4/RbQUqQnXhsYRHmmOT9GYiQ/view?embed' },
    { id: 11, title: 'คุณค่าของการเสียสละในอีดิลอัฎฮาต่อวิถีชีวิต', author: 'Chenasnin    Hajidereh', cover: 'https://www.canva.com/design/DAGtyH_VPwk/urOAXAsV6miokyGgLYypfA/view?embed', coverType: 'iframe', description: 'นำคุณค่าของการเสียสละจากวันอีดมาปรับใช้ เพื่อสร้างสรรค์วิถีชีวิตประจำวันที่ดีงามและมีความสุข', tags: ['เสียสละ', 'พัฒนาตนเอง'], rating: 5, popularity: 91, releaseDate: '2025-08-12', canvaLink: 'https://www.canva.com/design/DAGtyH_VPwk/urOAXAsV6miokyGgLYypfA/view?embed' },
    { id: 12, title: 'อีดิลอัฎฮา ความรัก ความศรัทธา และความภักดี', author: 'Daria        Awae', cover: 'https://www.canva.com/design/DAGtyHldUtU/tPVepLw4JA8VtWOVj2Twdw/view?embed', coverType: 'iframe', description: 'ดื่มด่ำกับเรื่องราวที่เปี่ยมไปด้วยความรัก ความศรัทธา และความภักดีต่อพระผู้เป็นเจ้าในวันอีด', tags: ['ศรัทธา', 'ความรัก'], rating: 5, popularity: 99, releaseDate: '2025-09-05', canvaLink: 'https://www.canva.com/design/DAGtyHldUtU/tPVepLw4JA8VtWOVj2Twdw/view?embed' }
  ];
  return defaults.filter(book => book.canvaLink && book.canvaLink.toString().trim() !== '');
}
