<script setup>
import { ref, computed } from 'vue'
import { storeToRefs } from 'pinia'
import { useImportStore } from '../stores/import'
import { useExportStore } from '../stores/export'

const importStore = useImportStore()
const store = useExportStore()
const { excelTitles, exportCode, checkedAll } = storeToRefs(store)

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
    store.excelTitles.forEach(element => {
      if (element.excelTitle === excelTitle.value) {
        check = false
      }
    })
    if (check) {
      store.addExcelTitle({
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
        document.getElementsByClassName('field')[0].style.display = 'none'
        document.getElementsByClassName('title')[0].style.borderBottom = 'none'
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
    store.deleteExcelTitle(val)
  }
}
const clear = () => {
  if (confirm('It will clear import and export data.\nMake sure you want to clear？')) {
    store.$reset()
    importStore.$reset()
    location.reload()
  }
}
const submit = () => {
  isSubmit.value = true
  const countfield = store.excelTitles.length
  if (countfield === 0) {
    document.getElementsByClassName('field')[0].style.display = 'block'
    document.getElementsByClassName('title')[0].style.borderBottom = '2px solid #dc3545'
  }
  const forms = document.querySelectorAll('.needs-validation')
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
    const Extension = ((store.fileExtension === 'xlsx') ? 'X' : 'H')
    let code = 'public void ExportExcel(DataTable dt)\n{\n' +
              '    MemoryStream MS = new MemoryStream();\n\n' +
              '    NPOI.SS.UserModel.IWorkbook book = BuildWorkbook(dt);\n\n' +
              '    string fileName = string.Format("{0:HHmmss}", DateTime.Now) + "' + store.fileName + '";\n\n' +
              '    HttpResponse httpResponse = HttpContext.Current.Response;\n' +
              '    httpResponse.Clear();\n' +
              '    httpResponse.Buffer = true;\n' +
              '    httpResponse.Charset = Encoding.UTF8.BodyName;\n' +
              '    httpResponse.AppendHeader("Content-Disposition", "attachment;filename=" + fileName + ".' + store.fileExtension + '");\n' +
              '    httpResponse.ContentEncoding = Encoding.UTF8;\n' +
              '    httpResponse.ContentType = "application/vnd.ms-excel; charset=UTF-8";\n' +
              '    book.Write(httpResponse.OutputStream);\n' +
              '    httpResponse.End();\n}\n' +
              'public ' + Extension + 'SSFWorkbook BuildWorkbook (DataTable dtData)\n{\n' +
              '    var book = new ' + Extension + 'SSFWorkbook();\n' +
              '    if (dtData.Rows.Count > 0)\n    {\n' +
              '        // font style\n' +
              '        ' + Extension + 'SSFFont font = (' + Extension + 'SSFFont)book.CreateFont();\n' +
              '        font.FontHeightInPoints = ' + store.fontSize + ';\n' +
              '        font.FontName = "' + store.fontFamily + '";\n\n' +
              '        // cell style\n' +
              '        ICellStyle CellStyle = book.CreateCellStyle();\n' +
              '        CellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;\n' +
              '        CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;\n' +
              '        CellStyle.WrapText = true;\n' +
              '        // cell border\n' +
              '        // CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;\n' +
              '        // CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;\n' +
              '        // CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;\n' +
              '        // CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;\n' +
              '        CellStyle.SetFont(font);\n\n' +
              '        ISheet sheet = book.CreateSheet("' + store.sheetName + '");\n' +
              '        IRow row0 = sheet.CreateRow(0);\n\n'
    const CellTitle = []
    const SqlColumn = []
    store.excelTitles.forEach(element => {
      if (element.checked) {
        CellTitle.push('"' + element.excelTitle + '"')
        SqlColumn.push('"' + element.columnName + '"')
      }
    })
    code += '        // excel title\n' +
            '        string[] TitleArr = { ' + CellTitle.join(', ') + ' };\n' +
            '        // sql column name\n' +
            '        string[] SqlDataArr = { ' + SqlColumn.join(', ') + ' };\n\n' +
            '        // setting cell title\n' +
            '        for (int i = 0; i < TitleArr.Length; i++)\n        {\n' +
            '           // create cell\n' +
            '           CreateICell(row0, CellStyle, i, TitleArr[i].ToString());\n' +
            '           // single cell width (10 is your width)\n' +
            '           sheet.SetColumnWidth(i, (int)((10 + 0.72) * 256));\n' +
            '           // auto width\n' +
            '           // sheet.AutoSizeColumn(i);\n        }\n\n' +
            '        // setting cell content\n' +
            '        for (int i = 0; i < dtData.Rows.Count; i++)\n        {\n' +
            '           // create row\n' +
            '           IRow row1 = sheet.CreateRow(i + 1);\n\n' +
            '           // setting row height\n' +
            '           // method 1\n' +
            '           // row1.HeightInPoints = 16;\n' +
            '           // method 2 (16 is your height)\n' +
            '           // row1.Height = 16 * 20;\n' +
            '           // method 3 (16 is your height)\n' +
            '           sheet.DefaultRowHeight = 16 * 20;\n\n' +
            '           for (int j = 0; j < SqlDataArr.Length; j++)\n           {\n' +
            '              CreateICell(row1, CellStyle, j, dtData.Rows[i][SqlDataArr[j]].ToString());\n           }\n        }\n    }\n' +
            '    else\n    {\n' +
            '       ISheet sheet = book.CreateSheet("' + store.sheetName + '");\n' +
            '       IRow drow0 = sheet.CreateRow(0);\n' +
            '       ICell cell0_0 = drow0.CreateCell(0, CellType.String);\n' +
            '       cell0_0.SetCellValue("");\n    }\n\n    return book;\n}\n' +
            'public void CreateICell (IRow row, ICellStyle CellStyle, int DataColumn, string ValueStr)\n{\n' +
            '    ICell cell0 = row.CreateCell((DataColumn), CellType.String);\n' +
            '    cell0.SetCellValue(ValueStr);\n' +
            '    cell0.CellStyle = CellStyle;\n}'
    store.exportCode = code
  }
}
const filePath = computed({
  get: () => store.filePath,
  set: (val) => {
    store.filePath = val
  }
})
const fileName = computed({
  get: () => store.fileName,
  set: (val) => {
    store.fileName = val
  }
})
const fontFamily = computed({
  get: () => store.fontFamily,
  set: (val) => {
    store.fontFamily = val
  }
})
const fontSize = computed({
  get: () => store.fontSize,
  set: (val) => {
    if (val < 0) val = 0
    store.fontSize = val
  }
})

const sheetName = computed({
  get: () => store.sheetName,
  set: (val) => {
    store.sheetName = val
  }
})

const fileExtension = computed({
  get: () => store.fileExtension,
  set: (val) => {
    store.fileExtension = val
  }
})

</script>

<template lang="pug">
.container
  .row
    .col-12
      form.row.g-3.needs-validation(novalidate)
        .col-md-4
          label.form-label(for='validationFileName') File Name
          input.form-control#validationFileName(type='text' v-model="fileName" required)
          .invalid-feedback Please provide a file name.

        .col-md-4
          label.form-label(for='validationFontFamily') Font Family
          input.form-control#validationFontFamily(type='text' v-model="fontFamily" required)
          .invalid-feedback Please provide a valid font family.

        .col-md-4
          label.form-label(for='validationFontSize') Font Size
          input.form-control#validationFontSize(type='number' v-model="fontSize" required min="0" )
          .invalid-feedback Please provide a valid font size.

        .col-md-4
          label.form-label(for='validationFontSize') Sheet Name
          input.form-control#validationFontSize(type='text' v-model="sheetName" required)
          .invalid-feedback Please provide a valid sheet name.

        .col-md-4
          label.form-label.mb-3 File Extension
          br
          input.form-check-input.me-2#validationExtensionXlsx(type='radio' name='radio-stacked' value="xlsx" required v-model="fileExtension")
          label.form-check-label.me-3(for='validationExtensionXlsx') .xlsx
          input.form-check-input.me-2#validationExtensionXls(type="radio" name="radio-stacked" value="xls" required v-model="fileExtension")
          label.form-check-label(for='validationExtensionXls') .xls
          .invalid-feedback Please provide a valid file extension.

        .col-md-12
          label.form-label Excel Cell
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
                    .field.invalid-feedback.field-error.text-start Please provide a valid excel cell.
              tr.title
                th
                th
                  input#check-all-col.form-check-input(type="checkbox" :checked="checkedAll" @click="store.checkAll()")
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
                    input.form-check-input(type="checkbox" :checked="element.checked" @click="store.checkItem(index)")
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
                    button.btn.btn-success.d-flex.me-2(type='button' @click="store.editExcelTitle(index)")
                      vue-feather(type="edit" stroke="white")
                    button.btn.btn-danger.d-flex(type='button' @click="deleteExcelTitle(index)")
                      vue-feather(type="trash-2" stroke="white")
                  td.d-flex.justify-content-center(v-else)
                    button.btn.btn-success.d-flex.me-2(type='button' @click="store.updateExcelTitle(index)")
                      vue-feather(type="check" stroke="white")
                    button.btn.btn-danger.d-flex(type='button' @click="store.cancelExcelTitle(index)")
                      vue-feather(type="x" stroke="white")
        .col-12.text-center.mb-3
          button.btn.btn-danger.me-5(type='button' @click="clear") Clear
          button.btn.btn-primary(type='button' @click="submit") Submit
        .col-12
          pre(v-highlightjs)
            code.javascript {{ exportCode }}
</template>
