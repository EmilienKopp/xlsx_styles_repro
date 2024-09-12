import XLSX from 'xlsx';
import fs from 'fs';

export class ExportService {
	data;
	workbook;

	static editAndSaveWorkbook(
		inputPath,
		outputPath,
		bookType
	) {
		const workbook = this.getWorkbookFromXlsFile(inputPath);

		// Edit the workbook
		const sheet = workbook.Sheets['観点別成績診断一覧（学習の様子あり）'];
		if (sheet) {
			sheet.A6 = { t: 's', v: '知・技', w: '知・技' };
		}

		// Save the workbook to a file
		const wopts = {
			bookType,
			type: 'buffer',
			cellStyles: true
		};
    
    const fileData = XLSX.writeXLSX(workbook,  wopts);
    fs.writeFileSync(outputPath, fileData);
	}

	static preserveStyles(
		sheet,
	) {
    console.log("Working on preserving header styles", sheet);
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for(let R = range.s.r; R <= range.e.r; ++R) {
      for(let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = sheet[address];
        if(cell && cell.s) {
          console.log(cell);
          if(cell.s.bgColor?.rgb == "FFFFFF" && cell.s.fgColor && cell.s.fgColor?.rgb != 'FFFFFF') {
            // set bg to be the same as fg
            cell.s.bgColor = cell.s.fgColor;
          }
        }
      }
    }

    console.log(sheet);
	}



	static getWorkbookFromXlsFile(path) {
		const data = fs.readFileSync(path);
		const wb = XLSX.read(data, { type: 'buffer', cellStyles: true, cellNF: true });
		return wb;
	}

}
