import ExcelJS from "exceljs";
import sharp from "sharp";
import { existsSync, mkdirSync, writeFileSync } from "fs";
import { join } from "path";

//yyyy-mm-dd 로 입력하거나 yyyy-mm 으로 입력한 경우가 있어 분기처리
function validDate(cellValue) {
  if (typeof cellValue === "string") {
    return cellValue; // 원래 문자열 그대로 반환
  }

  if (cellValue instanceof Date) {
    const date = new Date(cellValue);
    return isNaN(date.getTime()) ? "" : date.toLocaleDateString("en-CA");
  }

  return "";
}

async function parseExcelFile() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("./original.xlsx"); // 엑셀 파일 경로

  const ACTIVYTY_ROW_INDEX = 9;

  const images = [];
  const results = {};

  workbook.eachSheet((sheet, sheetId) => {
    if (sheet.name === "학교목록") {
      return;
    }

    const schoolName = sheet.getCell("B2").text;
    const schoolCode = sheet.getCell("C2").text;

    const groupName = sheet.getCell("B3").text;
    const groupcode = sheet.getCell("C3").text;

    if (!schoolName || !groupcode) {
      return;
    }

    if (!images[sheetId]) {
      images[sheetId] = {}; // 객체 초기화
    }

    const folderPath = join("riu", "images", schoolCode, groupcode);

    // 폴더가 존재하지 않으면 생성 (중간 폴더도 자동으로 생성)
    if (!existsSync(folderPath)) {
      mkdirSync(folderPath, { recursive: true });
    }
    // 이미지 로컬 파일 생성
    const sheetImages = sheet.getImages();
    sheetImages.forEach((image, index) => {
      const imageData = workbook.model.media.find(
        (media) => media.index === image.imageId
      );

      if (imageData && imageData.buffer) {
        let buffer = imageData.buffer;

        // 이미지 압축
        if (image.range.ext.width > 300) {
          sharp(buffer)
            .resize({ width: 300 }) // riu 웹사이트 이미지 사이즈 : 가로 최대 300px
            .jpeg({ quality: 70 })
            .toBuffer()
            .then((resizedBuffer) => {
              buffer = resizedBuffer;
            });
        } else {
          sharp(buffer)
            .jpeg({ quality: 70 })
            .toBuffer()
            .then((resizedBuffer) => {
              buffer = resizedBuffer;
            });
        }

        // 이미지 파일로 저장
        const filePath = join(folderPath, `image_${index}.jpg`);
        writeFileSync(filePath, buffer);

        // 이미지 path 정보 생성
        const position = sheet.getCell(
          image.range.tl.nativeRow + 1,
          image.range.tl.nativeCol + 1
        ).address;

        images[sheetId] = {
          ...(images[sheetId] && images[sheetId]),
          [position]: filePath,
        };
      }
    });

    // 동아리 정보 생성
    const establishedAt = sheet.getCell("B4").value;
    const disbandedAt = sheet.getCell("C4").value;
    const snsLink = sheet.getCell("B5").text;

    if (!results[schoolCode]) {
      results[schoolCode] = {
        name: schoolName, // 객체 초기화
        code: schoolCode,
        groups: {},
      };
    }

    if (!results[schoolCode].groups[groupcode]) {
      results[schoolCode].groups[groupcode] = {
        name: groupName, // 객체 초기화
        code: groupcode,
        establishedAt: validDate(establishedAt),
        disbandedAt: validDate(disbandedAt),
        snsLink,
        logo: images[sheetId] && images[sheetId]["E2"],
        activities: [],
      };
    }

    // 활동 내역 생성
    for (let row = ACTIVYTY_ROW_INDEX; row <= sheet.rowCount; row++) {
      const activityTitle = sheet.getCell(`D${row}`).text;

      if (!activityTitle) {
        continue;
      }

      const activityYear = sheet.getCell(`A${row}`).text;
      const activitySeason = sheet.getCell(`B${row}`).text;
      const activityPeriod = sheet.getCell(`C${row}`).text;
      const activityDetails = sheet.getCell(`E${row}`).text;
      const activityExtraLink = sheet.getCell(`F${row}`).text;

      results[schoolCode].groups[groupcode].activities.push({
        year: activityYear,
        season: activitySeason,
        period: activityPeriod,
        title: activityTitle,
        details: activityDetails,
        extraLink: activityExtraLink,
        image: images[sheetId] && images[sheetId][`G${row}`],
      });
    }
  });

  // JSON 파일 저장
  const outputDir = "riu/output";
  if (!existsSync(outputDir)) {
    mkdirSync(outputDir, { recursive: true });
  }

  const outputFilePath = join(outputDir, "results.json");
  writeFileSync(outputFilePath, JSON.stringify(results, null, 2), "utf-8");

  console.log("Done!");

  return results;
}

parseExcelFile();

export default parseExcelFile;
