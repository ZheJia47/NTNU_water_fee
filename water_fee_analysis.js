var XLSX = require('XLSX')

// Transform_excel_file ####################################################
function Transform_excel_file( input_file, 
  output_file_name = input_file.substring(0, input_file.lastIndexOf('.')) + '.xlsx' ){
  // read file 
  var input_file = XLSX.readFile(input_file)
  // Get worksheet 
  var input_sheet = input_file.Sheets[input_file.SheetNames[0]]  
  // var desired_value = (desired_cell ? desired_cell.v : undefined)

  // create a new workbook ##########################################
  var output_file = XLSX.utils.book_new();
  var worksheet = {};   

  // worksheet range
  var range = {s:{r: 0, c: 0},
               e: {r: 40, c: 20}}
               
  worksheet['!ref'] = XLSX.utils.encode_range(range)      
               
  // todo: borderAll, column A 加寬
  const borderAll = {  //单元格外侧框线
    top: {
      style: 'thin',
      color: { auto: 1 } 
    },
    bottom: {
      style: 'thin'
    },
    left: {
      style: 'thin'
    },
    right: {
      style: 'thin'
    }
  }
  
  // Cell_cnt ###########################################################
  function Cell_cnt(cnt){
    cnt = {v: String(cnt), 
           s: {alignment: {horizontal: 'center'},
               border: borderAll
              }
          }
    return cnt
  }

  // worksheet content
  worksheet['A1'] = Cell_cnt('日期 (date)')
  worksheet['A2'] = Cell_cnt('寢室號碼\n(room number)')

  // 2nd row content: floor number
  for(var ifloor=2; ifloor<=9; ifloor++)
    worksheet[XLSX.utils.encode_cell({r: 2, c: ifloor-1})] 
    = Cell_cnt('70' + String(ifloor) + 'F')
  for(var ifloor=10; ifloor<=12; ifloor++)
    worksheet[XLSX.utils.encode_cell({r: 2, c: ifloor-1})] 
    = Cell_cnt('7' + String(ifloor) + 'F')

  // 1st column content: room number
  for(var inumber=1; inumber<=30; inumber++)
    worksheet[XLSX.utils.encode_cell({r: inumber+2, c: 0})] 
    = Cell_cnt(inumber)

  // water fee ########################################################
  var fee = [null,null]
  
  for(var icolumn=1; icolumn<=11; icolumn++){
    fee.push([]); fee[icolumn+1].push(null)
    for(var i_input_row=3; i_input_row<=252; i_input_row++){
      for(var inumber=1; inumber<=30; inumber++){

        if(String(input_sheet[XLSX.utils.encode_cell({r: i_input_row, c: 1})].v).slice(0,3) 
            == String(worksheet[XLSX.utils.encode_cell({r: 2, c: icolumn})].v).slice(0,3)){

          if(Number(String(input_sheet[XLSX.utils.encode_cell({r: i_input_row, c: 1})].v).slice(3,5)) 
              == Number(inumber)){    
          
            worksheet[XLSX.utils.encode_cell({r: inumber+2, c:icolumn})] 
            = Cell_cnt(input_sheet[XLSX.utils.encode_cell({r: i_input_row, c: 5})].v)

            fee[icolumn+1].push(input_sheet[XLSX.utils.encode_cell({r: i_input_row, c: 5})].v)

          }
          
        }          
      }
    }
  }

  // 各樓層平均值 (不包括0) ######################################
  worksheet[XLSX.utils.encode_cell({r: 34, c: 0})] = Cell_cnt('各樓平均值 (不包括0)')

  function floor_average(floor_fee){
    var floor_fee_new = []

    for(i=1; i<floor_fee.length; i++)
      if(floor_fee[i]>0)
        floor_fee_new.push(floor_fee[i])

    var total = 0
    for(i=0; i<floor_fee_new.length; i++)
      total += floor_fee_new[i]

    average = total/floor_fee_new.length
    return average
  }

  for(var icolumn=1; icolumn<=11; icolumn++)
    worksheet[XLSX.utils.encode_cell({r: 35, c: icolumn})] = Cell_cnt(floor_average(fee[icolumn+1]).toFixed(2))

  // 男女研究生及大學生樓層各別平均值 (不包括0) ###########################
  // 大學生宿舍4人房
  // 研究生宿舍3人房
  // 2F: male graduate student
  // 3~5F: male under graduate student
  // 6~9F: female under graduate student
  // 10~12F: female graduate student  
  // average for each case
  function case_average(begin_floor, end_floor){
    var total = 0
    for(var ifloor=begin_floor; ifloor<=end_floor; ifloor++)
      total += floor_average(fee[ifloor])
    average = total/(end_floor-begin_floor+1)

    return average    
  }

  var male_grad_avg = case_average(2,2)
  var male_under_grad_avg = case_average(3,5)
  var female_grad_avg = case_average(10,12)
  var female_under_grad_avg = case_average(6,9)
  
  // 補缺失值
  for(var icolumn=1; icolumn<=11; icolumn++)
    for(var inumber=1; inumber<=30; inumber++)
      if(worksheet[XLSX.utils.encode_cell({r: inumber+2, c: icolumn})] == undefined || worksheet[XLSX.utils.encode_cell({r: inumber+2, c: icolumn})].v <= 0){
        
        if(icolumn==1) // male graduate student case
          worksheet[XLSX.utils.encode_cell({r: inumber+2, c: icolumn})] = Cell_cnt(male_grad_avg.toFixed(2))

        if(icolumn>=2 && icolumn<=4) // male under graduate student case
          worksheet[XLSX.utils.encode_cell({r: inumber+2, c: icolumn})] = Cell_cnt(male_under_grad_avg.toFixed(2))

        if(icolumn>=5 && icolumn<=8) // female graduate student case
          worksheet[XLSX.utils.encode_cell({r: inumber+2, c: icolumn})] = Cell_cnt(female_under_grad_avg.toFixed(2))

        if(icolumn>=9 && icolumn<=11) // female under graduate student case
          worksheet[XLSX.utils.encode_cell({r: inumber+2, c: icolumn})] = Cell_cnt(female_grad_avg.toFixed(2))

      }
  
  // output file
  output_file.SheetNames.push('push')
  output_file.Sheets.push = worksheet
  XLSX.writeFile(output_file, output_file_name)
}

// excel to html table ####################################################
var fs = require('fs')

function Excel_to_html_table(input_file, 
  output_file_name = 'html_table/'+input_file.substring(0, input_file.lastIndexOf('.')) + '.html'){
  // read file 
  var input_file_name = input_file
  var input_file = XLSX.readFile(input_file)
  // Get worksheet 
  var input_sheet = input_file.Sheets[input_file.SheetNames[0]]  
  // new html file
  fs.open(output_file_name,'w',function(){})
  var cnt = '<!DOCTYPE html> \n'
  cnt += '<html> \n\
  <head> \n\
    <meta charset="utf-8"> \n\
    <title>NTNU water fee</title> \n\
  </head> \n\n\
  <body> \n'
  
  // 1st row cnt
  cnt += '<table border="1"> \n\
  <thead> \n\
    <tr>\n      <th>寢室號碼</th> \n'
    for(var ifloor=2; ifloor<=9; ifloor++)
      cnt += '      <th>70'+String(ifloor)+'F</th> \n'
    for(var ifloor=10; ifloor<=12; ifloor++)
      cnt += '      <th>7'+String(ifloor)+'F</th> \n'
  cnt += '    </tr> \n'
  cnt += '  </thead> \n\n'    
  // tbody
  cnt += '  <tbody> \n'
  for(var inumber=1; inumber<=30; inumber++){ //row
    cnt += '    <tr> \n'
    for(var ifloor=1; ifloor<=12; ifloor++){ //column
      if(ifloor==1) cnt += '      <td>'+String(inumber)+'</td> \n'
      else cnt += '      <td>'+input_sheet[XLSX.utils.encode_cell({r: inumber+2, c: ifloor-1})].v+'</td> \n'
    }
    cnt += '    </tr> \n'
  }
  cnt += '  </tbody> \n</table>\n'
  cnt += '  <a href="../'+String(input_file_name)+'" download><p>download data</p></a>\n'
  cnt += '  </body> \n</html>'
  fs.writeFile(output_file_name, cnt,function(){})
}

// excute function ###########################################################
Excel_to_html_table('2020-09-17至2020-09-17全部房號的用水記錄報表.xlsx')





// ref. code 
function ExcelExport(rows, cols, fileName) {
  var ws
  ws = XLSX.utils.json_to_sheet(rows)
  const ec = (r, c) => {
    return XLSX.utils.encode_cell({ r: r, c: c })
  }
  const deleteCol = (ws, colIndex) => {
    let range = XLSX.utils.decode_range(ws['!ref'])
    for (var R = range.s.c; R <= range.e.r; ++R) {
      for (var C = colIndex; C <= range.e.c; ++C) {
        ws[ec(R, C)] = ws[ec(R, C + 1)]
      }
    }
    range.e.c--
    ws['!ref'] = XLSX.utils.encode_range(range.s, range.e)
  }
  var col2Idx = 0
  let range = XLSX.utils.decode_range(ws['!ref'])
  for (var l = range.e.c; l > -1; l--) {
    for (col2Idx = 0; col2Idx < cols.length; col2Idx++) {
      if (cols[col2Idx].name === ws[ec(0, l)].v) {
        break
      }
    }
    if (col2Idx >= cols.length) {
      deleteCol(ws, l)
    }
  }
  range = XLSX.utils.decode_range(ws['!ref'])
  for (l = 0; l <= (range.e.c); l++) {
    cols.forEach(col2 => {
      if (col2.name === ws[ec(0, l)].v) {
        ws[ec(0, l)].v = col2.label
      }
    })
  }
  var wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, '表單')
  XLSX.writeFile(wb, fileName + '.xlsx')
}