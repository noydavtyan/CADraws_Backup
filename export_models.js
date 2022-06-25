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
  year + "-" + month + "-" + date + "_" + hours + "-" + minutes + "-" + seconds;

const fetchRows = () => {
  axios
    .request({
      url: "http://localhost:8095/models",
      method: "GET",
      headers: {
        Authorization:
          "secretPrefix eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJhIiwiaWF0IjoxNjM4MTMyMjYwLCJleHAiOjE2MzgxMzIzNjB9.H_aeOjNBkKJOV3raman5odIoDSRiJDnXxS7LwsodPnZCRRuGIrOnGESyRCzaZD0oy_O1XOZnsTTKx3g2iblbgw",
      },
    })
    .then((response) => {
      worksheet.cell(i, 1).string("ID");
      worksheet.cell(i, 2).string("Company");
      worksheet.cell(i, 3).string("Description");
      worksheet.cell(i, 4).string("Order Date");
      worksheet.cell(i, 5).string("Delivery Date");
      worksheet.cell(i, 6).string("Worker Price");
      worksheet.cell(i, 7).string("Total Price");
      worksheet.cell(i, 8).string("Payed");
      worksheet.cell(i, 9).string("Received");
      response.data.forEach((model) => {
        i++;
        worksheet.cell(i, 1).string(model.modelId);
        worksheet.cell(i, 2).string(model.company.name);
        worksheet.cell(i, 3).string(model.description);
        worksheet.cell(i, 4).string(model.worker_price);
        worksheet.cell(i, 5).string(model.total_price);
        worksheet.cell(i, 6).string(model.orderDate);
        worksheet.cell(i, 7).string(model.delivery_date);
        worksheet.cell(i, 8).string(model.payed);
        worksheet.cell(i, 9).string(model.received);
      });
      workbook.write(fileName + ".xlsx");
    })
    .catch((error) => {
      console.log(error);
    });
};

fetchRows();
