const axios = require("axios");
const excel = require("excel4node");

var workbook = new excel.Workbook();
let worksheet = workbook.addWorksheet("Models");
let i = 1;
let date_ob = new Date();

const date = ("0" + date_ob.getDate()).slice(-2);
const month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
const year = date_ob.getFullYear();
const hours = date_ob.getHours();
const minutes = date_ob.getMinutes();
const seconds = date_ob.getSeconds();
const fileName =
  "companies_" +
  year +
  "-" +
  month +
  "-" +
  date +
  "_" +
  hours +
  "-" +
  minutes +
  "-" +
  seconds;

const fetchRows = () => {
  axios
    .request({
      url: "http://localhost:8095/allClients",
      method: "GET",
      headers: {
        Authorization:
          "secretPrefix eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJhIiwiaWF0IjoxNjM4MTMyMjYwLCJleHAiOjE2MzgxMzIzNjB9.H_aeOjNBkKJOV3raman5odIoDSRiJDnXxS7LwsodPnZCRRuGIrOnGESyRCzaZD0oy_O1XOZnsTTKx3g2iblbgw",
      },
    })
    .then((response) => {
      worksheet.cell(i, 1).string("ID");
      worksheet.cell(i, 2).string("alreadyClient");
      worksheet.cell(i, 3).string("city");
      worksheet.cell(i, 4).string("contact_person");
      worksheet.cell(i, 5).string("country");
      worksheet.cell(i, 6).string("dateAdded");
      worksheet.cell(i, 7).string("email");
      worksheet.cell(i, 8).string("is_sent");
      worksheet.cell(i, 9).string("linkedin");
      worksheet.cell(i, 10).string("name");
      worksheet.cell(i, 11).string("owner");
      worksheet.cell(i, 12).string("plz");
      worksheet.cell(i, 13).string("refer");
      worksheet.cell(i, 14).string("street");
      worksheet.cell(i, 15).string("website");
      worksheet.cell(i, 16).string("xing");
      worksheet.cell(i, 17).string("representative");
      response.data.forEach((company) => {
        i++;
        worksheet.cell(i, 1).number(company.id);
        worksheet.cell(i, 2).string(company.alreadyClient);
        if (company.city) worksheet.cell(i, 3).string(company.city);
        if (company.contact_person)
          worksheet.cell(i, 4).string(company.contact_person);
        worksheet.cell(i, 5).string(company.country);
        worksheet.cell(i, 6).string(company.dateAdded);
        worksheet.cell(i, 7).string(company.email);
        worksheet.cell(i, 8).number(company.isSent);
        if (company.linkedin) worksheet.cell(i, 9).string(company.linkedin);
        if (company.name) worksheet.cell(i, 10).string(company.name);
        if (company.owner) worksheet.cell(i, 11).string(company.owner);
        if (company.plz) worksheet.cell(i, 12).string(company.plz);
        if (company.refer) worksheet.cell(i, 13).string(company.refer);
        if (company.street) worksheet.cell(i, 14).string(company.street);
        if (company.website) worksheet.cell(i, 15).string(company.website);
        if (company.xing) worksheet.cell(i, 16).string(company.xing);
        if (company.representative)
          worksheet.cell(i, 17).string(company.representative);
      });
      workbook.write(fileName + ".xlsx");
    })
    .catch((error) => {
      console.log(error);
    });
};

fetchRows();
