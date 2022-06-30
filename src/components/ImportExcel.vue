<script setup>
import { ref, computed } from 'vue'
import { storeToRefs } from 'pinia'
import { useExportStore } from '../stores/import'
const store = useExportStore()
const { excelTitles, exportCode, checkedAll } = storeToRefs(store)

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
    let code = ''
    for (let i = 65; i <= 90; i++) {
      code += String.fromCharCode(i)
    }
    store.exportCode = code
  }
}
</script>

<template lang="pug">
.container
  .row
    .col-12
      form.row.g-3.needs-validation(novalidate)
        .col-md-4
          label.form-label(for='validationFilePath') File Path
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

        .col-md-12
          label.form-label Excel Cell
          table#cellTable.table.table-striped
            thead
              tr
                td.pt-0(colspan="5")
                  .form-group.row.d-flex.align-items-center
                    input.form-control.no-validate.me-2(type='text' style="width:40%;" v-model="excelTitle" placeholder="Excel Title")
                    input.form-control.no-validate.me-2(type='text' style="width:40%;" v-model="sqlColumnName" placeholder="SQL Column Name")
                    button.btn.btn-primary.d-flex(type='button' @click="addExcelTitle" style="width:auto")
                      vue-feather(type="plus" stroke="white")
                    .field.invalid-feedback.field-error.text-start Please provide a valid excel cell.
              tr.title
                th
                th
                  input#check-all-col.form-check-input(type="checkbox")
                th Excel Title
                th SQL Column Name
                th Actions
            draggable.drag(v-model="excelTitles" tag="tbody" item-key="excelTitle" handle=".handle" :animation="100")
              template(#item="{ element, index }")
                tr
                  td
                    vue-feather.handle(type="align-justify" stroke="black")
                  td
                    input.form-check-input(type="checkbox")
                  td(v-if="!element.edit") {{ element.excelTitle }}
                  td(v-else)
                    input.form-control.me-2(type='text' v-model="element.editExcelTitle")
                  td(v-if="!element.edit") {{ element.columnName }}
                  td(v-else)
                    input.form-control.me-2(type='text' v-model="element.editColumnName")
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
          button.btn.btn-danger.me-5(type='button') Clear
          button.btn.btn-primary(type='button' @click="submit") Submit
        .col-12
          pre(v-highlightjs)
            code.javascript {{ exportCode }}
</template>
