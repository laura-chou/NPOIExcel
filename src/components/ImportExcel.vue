<script setup>
import { ref, computed } from 'vue'
import { storeToRefs } from 'pinia'
import { useImportStore } from '../stores/import'
import { useExportStore } from '../stores/export'
const store = useImportStore()
const exportStore = useExportStore()
const { importCode } = storeToRefs(store)
const { excelTitles, checkedAll } = storeToRefs(exportStore)

const excelTitle = ref('')
const sqlColumnName = ref('')
const sqlColumnType = ref('')
const isSubmit = ref(false)

const addExcelTitle = () => {
  if (excelTitle.value === '') {
    alert('Please input excel title.')
  } else if (sqlColumnName.value === '') {
    alert('Please input sql column name.')
  } else if (sqlColumnType.value === '') {
    alert('Please select type.')
  } else {
    let check = true
    exportStore.excelTitles.forEach(element => {
      if (element.excelTitle === excelTitle.value) {
        check = false
      }
    })
    if (check) {
      exportStore.addExcelTitle({
        excelTitle: excelTitle.value,
        columnName: sqlColumnName.value,
        columnType: sqlColumnType.value,
        edit: false,
        editExcelTitle: excelTitle.value,
        editColumnName: sqlColumnName.value,
        editColumnType: sqlColumnType.value,
        checked: true
      })
      if (isSubmit.value) {
        document.getElementsByClassName('field-import')[0].style.display = 'none'
        document.getElementsByClassName('title-import')[0].style.borderBottom = 'none'
      }
    } else {
      alert('Duplicate excel title.')
    }
    excelTitle.value = ''
    sqlColumnName.value = ''
    sqlColumnType.value = ''
  }
}
const deleteExcelTitle = (val) => {
  if (confirm('Make sure you want to delete？')) {
    exportStore.deleteExcelTitle(val)
  }
}

const clear = () => {
  if (confirm('It will clear import and export data.\nMake sure you want to clear？')) {
    store.$reset()
    exportStore.$reset()
    location.reload()
  }
}

const submit = () => {
  isSubmit.value = true
  const countfield = exportStore.excelTitles.length
  if (countfield === 0) {
    document.getElementsByClassName('field-import')[0].style.display = 'block'
    document.getElementsByClassName('title-import')[0].style.borderBottom = '2px solid #dc3545'
  }
  const forms = document.querySelectorAll('.needs-validation-import')
  let validate = false
  Array.prototype.slice.call(forms)
    .forEach(function (form) {
      if (!form.checkValidity()) {
        form.classList.add('was-validated')
      } else if (countfield === 0) {
        form.classList.add('was-validated')
      } else {
        validate = true
      }
    })
  if (validate) {
    let code = 'public string UpLoadFile (FileUpload UploadFiles)\n{\n'
    code += '  string returnMsg = "";\n\n'
    code += '  if (UploadFiles.HasFile)\n  {\n'
    code += '    string fileName = UploadFiles.FileName;\n'
    code += '    string extension = Path.GetExtension(fileName).ToLowerInvariant();\n'
    code += '    // upload file limits\n'
    code += '    string[] allowedExtextsion = { ".xlsx", ".xls" };\n'
    code += '    if (!allowedExtextsion.Contains(extension))\n    {\n'
    code += '      returnMsg = "Only accept Excel file.";\n    }\n    else\n    {\n'
    code += '      // file size limits\n'
    code += '      int filesize = UploadFiles.PostedFile.ContentLength;\n'
    code += '      if (filesize > ' + store.fileSizeLimit + ')\n      {\n'
    code += '        returnMsg = "File too large.";\n      }\n      else\n      {\n'
    code += '        // file storage path\n'
    code += '        string serverDir = Server.MapPath("~/' + store.filePath + '");\n'
    code += '        if (!Directory.Exists(serverDir)) Directory.CreateDirectory(serverDir);\n\n'
    code += '        // rename if the file name is duplicate\n'
    code += '        string serverFilePath = Path.Combine(serverDir, fileName);\n'
    code += '        string fileNameOnly = Path.GetFileNameWithoutExtension(fileName);\n'
    code += '        int count = 1;\n'
    code += '        while (File.Exists(serverFilePath))\n        {\n'
    code += '          fileName = string.Concat(fileNameOnly, "_", count, extension);\n'
    code += '          serverFilePath = Path.Combine(serverDir, fileName);\n'
    code += '          count++;\n        }\n\n'
    code += '        // file storage\n'
    code += '        UploadFiles.SaveAs(serverFilePath);\n\n'
    code += '        // read file\n'
    code += '        IWorkbook wk;\n'
    code += '        ISheet hst;\n\n'
    code += '        using (FileStream fs = new FileStream(serverFilePath, FileMode.Open, FileAccess.ReadWrite))\n        {\n'
    code += '          if (Path.GetExtension(fileName).Equals(".xls"))\n          {\n'
    code += '             wk = new HSSFWorkbook(fs);\n          }\n          else\n          {\n'
    code += '             wk = new XSSFWorkbook(fs);\n          }\n'
    code += '        }\n\n'
    code += '        hst = (ISheet)wk.GetSheetAt(0);\n'
    code += '        returnMsg = GetCellData(hst);\n'
    if (store.saveFile === 'true') code += '        // delete file\n        System.IO.File.Delete(serverFilePath);\n'
    code += '        GC.Collect();\n      }\n    }\n  }\n  else\n  {\n'
    code += '    returnMsg = "Please upload Excel file.";\n  }\n\n'
    code += '  return returnMsg;\n}\n'
    code += 'private string GetCellData (ISheet sheet)\n{\n'
    code += '  IRow rowTitle = (IRow)sheet.GetRow(0);\n'
    code += '  string message = "";\n'
    code += '  string errorNum = "";\n'
    const CellTitle = []
    const SqlColumn = []
    const ColumnType = []
    exportStore.excelTitles.forEach(element => {
      if (element.checked) {
        CellTitle.push('"' + element.excelTitle + '"')
        SqlColumn.push('"' + element.columnName + '"')
        ColumnType.push('' + element.columnType + '')
      }
    })
    code += '  string[] title = {' + CellTitle.join(', ') + '};\n'
    const ExcelNumber = []
    let count = CellTitle.length
    let runTime = 1
    let str = ''
    if (count > 26 && count < 53) {
      runTime++
    }
    if (count > 52 && count < 79) {
      runTime++
    }
    if (count > 78 && count < 105) {
      runTime++
    }
    for (let i = 0; i <= runTime; i++) {
      switch (i) {
        case 1:
          str = 'A'
          break
        case 2:
          str = 'B'
          break
        case 3:
          str = 'C'
          break
      }
      for (let j = 65; j <= 90; j++) {
        if (count !== 0) {
          ExcelNumber.push('"' + str + String.fromCharCode(j) + '1"')
          count--
        }
      }
    }

    code += '  string[] cellNumber = {' + ExcelNumber.join(', ') + '};\n\n'
    code += '  // confirm the number of titles\n'
    code += '  if (rowTitle.LastCellNum != title.Length)\n  {\n'
    code += '    return "Wrong number of title fields.";\n  }\n  else\n  {\n'
    code += '    // confirm title\n'
    code += '    for (int i = 0; i < title.Length; i++)\n    {\n'
    code += '      if (rowTitle.GetCell(i).ToString().Trim() != title[i])\n      {\n'
    code += '        message += (message != "" ? "、" : "") + cellNumber[i];\n      }\n    }\n'
    code += '    if (message != "") return message + " wrong title";\n  }\n\n'
    code += '  // create DataTable\n'
    code += '  DataTable dt = new DataTable();\n'
    code += '  string[] sqlColumn = {' + SqlColumn.join(', ') + '};\n'
    code += '  for (int i = 0; i < sqlColumn.Length; i++)\n  {\n'
    code += '    DataColumn dc = new DataColumn();\n'
    code += '    dc.DataType = Type.GetType("System.String");\n'
    code += '    dc.ColumnName = sqlColumn[i];\n'
    code += '    dt.Columns.Add(dc);\n  }\n\n'
    code += '  // read  the cell value and add it to the DataTable\n'
    code += '  for (int i = 1; i <= sheet.LastRowNum; i++)\n  {\n'
    code += '      IRow row = sheet.GetRow(i);\n'
    code += '      DataRow dr = dt.NewRow();\n'
    code += '      int itemNumber = i + 1;\n'
    code += '\n      for (int j = 0; j < title.Length; j++)\n      {\n'
    code += '          string value = row.GetCell(j).ToString().Trim();\n'
    code += '          // you can confirm or change the values,and then add them to the DataTable here\n'
    code += '          switch (j)\n          {\n'
    ColumnType.forEach((element, index) => {
      if (element !== 'Text') {
        code += '            case ' + index + ':\n'
      }
    })
    if (ColumnType.includes('Number')) {
      code += '              if (!value.All(char.IsDigit))\n              {\n'
      code += '                errorNum += ((errorNum == "") ? "" : "、") + cellNumber[j].Replace("1", "") + itemNumber.ToString();\n              }\n              else\n              {\n'
      code += '                dr[sqlColumn[j]] = value;\n              }\n'
      code += '              break;\n'
    }
    code += '            default:\n'
    code += '              dr[sqlColumn[j]] = value;\n'
    code += '              break;\n'
    code += '          }\n      }\n\n'
    code += '      dt.Rows.Add(dr);\n  }\n\n'
    code += '  if (errorNum != "")\n  {\n'
    code += '    errorNum = "Cell " + errorNum + ".Please input the number.";\n'
    code += '    message += errorNum;\n  }\n\n'
    code += '  if (message == "")\n  {\n'
    code += '    foreach (DataRow dr in dt.Rows)\n    {\n'
    code += '      // you can read the data and save it in the database here\n'
    code += '      // for example: dr[' + SqlColumn[0] + '].ToString();\n    }\n  }\n\n'
    code += '  return message;\n'
    code += '}'
    store.importCode = code
  }
}

const filePath = computed({
  get: () => store.filePath,
  set: (val) => {
    store.filePath = val
  }
})
const fileSizeLimit = computed({
  get: () => store.fileSizeLimit,
  set: (val) => {
    store.fileSizeLimit = val
  }
})
const saveFile = computed({
  get: () => store.saveFile,
  set: (val) => {
    store.saveFile = val
  }
})

</script>

<template lang="pug">
.container
  .row
    .col-12
      form.row.g-3.needs-validation-import(novalidate)
        .col-md-4
          label.form-label(for='validationFilePath') Import File Path
          input.form-control#validationFilePath(type='text' v-model="filePath" required)
          .invalid-feedback Please provide a valid file path.

        .col-md-4
          label.form-label Save File
          br
          input.form-check-input.me-2#validationSaveFileYes(type='radio' name='radio-stacked' value="true" required v-model="saveFile")
          label.form-check-label.me-3(for='validationSaveFileYes') Yes
          input.form-check-input.me-2#validationSaveFileNo(type="radio" name="radio-stacked" value="false" required v-model="saveFile")
          label.form-check-label(for='validationSaveFileNo') No
          .invalid-feedback Please provide a valid save file.

        .col-md-4
          label.form-label(for='validationFileSize') File Size Limit
          select#validationFileSize.form-select.me-2(v-model="fileSizeLimit" required)
            option(value="") Choose Size ...
            option(value="209715200") 200 MB
            option(value="524288000") 500 MB
            option(value="1073741824") 1 GB
            option(value="2147483648") 2 GB
            option(value="3221225472") 3 GB
          .invalid-feedback Please provide a file size limit.

        .col-md-12
          label.form-label.me-3 Excel Cell
          table#cellTable.table.table-striped
            thead
              tr
                td.pt-0(colspan="6")
                  .form-group.row.d-flex.align-items-center
                    input.form-control.no-validate.me-2(type='text' style="width:26%;" v-model="excelTitle" placeholder="Excel Title")
                    input.form-control.no-validate.me-2(type='text' style="width:26%;" v-model="sqlColumnName" placeholder="SQL Column Name")
                    select.form-select.no-validate.me-2(style="width:26%;" v-model="sqlColumnType")
                      option(value="") Choose Type...
                      option(value="Text") Text
                      option(value="Number") Number
                    button.btn.btn-primary.d-flex(type='button' @click="addExcelTitle" style="width:auto")
                      vue-feather(type="plus" stroke="white")
                    .field-import.invalid-feedback.field-error.text-start Please provide a valid excel cell.
              tr.title-import
                th
                th
                  input#check-all-col.form-check-input(type="checkbox" :checked="checkedAll" @click="exportStore.checkAll()")
                th Excel Title
                th SQL Column Name
                th Column Type
                th Actions
            draggable.drag(v-model="excelTitles" tag="tbody" item-key="excelTitle" handle=".handle" :animation="100")
              template(#item="{ element, index }")
                tr
                  td
                    vue-feather.handle(type="align-justify" stroke="black")
                  td
                    input.form-check-input(type="checkbox" :checked="element.checked" @click="exportStore.checkItem(index)")
                  td(v-if="!element.edit") {{ element.excelTitle }}
                  td(v-else)
                    input.form-control.me-2(type='text' v-model="element.editExcelTitle")
                  td(v-if="!element.edit") {{ element.columnName }}
                  td(v-else)
                    input.form-control.me-2(type='text' v-model="element.editColumnName")
                  td(v-if="!element.edit") {{ element.columnType }}
                  td(v-else)
                    select.form-select.me-2(v-model="element.editColumnType")
                      option(value="Text") Text
                      option(value="Number") Number
                  td.d-flex.justify-content-center(v-if="!element.edit")
                    button.btn.btn-success.d-flex.me-2(type='button' @click="exportStore.editExcelTitle(index)")
                      vue-feather(type="edit" stroke="white")
                    button.btn.btn-danger.d-flex(type='button' @click="deleteExcelTitle(index)")
                      vue-feather(type="trash-2" stroke="white")
                  td.d-flex.justify-content-center(v-else)
                    button.btn.btn-success.d-flex.me-2(type='button' @click="exportStore.updateExcelTitle(index)")
                      vue-feather(type="check" stroke="white")
                    button.btn.btn-danger.d-flex(type='button' @click="exportStore.cancelExcelTitle(index)")
                      vue-feather(type="x" stroke="white")
        .col-12.text-center.mb-3
          button.btn.btn-danger.me-5(type='button' @click="clear") Clear
          button.btn.btn-primary(type='button' @click="submit") Submit
        .col-12
          pre(v-highlightjs)
            code.javascript {{ importCode }}
</template>
