import { defineStore } from 'pinia'

export const useExportStore = defineStore({
  id: 'import',
  state: () => ({
    excelTitles: [],
    fileName: '',
    fontFamily: '',
    fontSize: '',
    sheetName: '',
    exportCode: '',
    checkedAll: true
  }),
  getters: {
    // columnNameArr: (state) => state.columnNameArr
  },
  actions: {
    addExcelTitle (value) {
      this.excelTitles.push(value)
    },
    updateExcelTitle (index) {
      this.excelTitles[index] = {
        excelTitle: this.excelTitles[index].editExcelTitle,
        columnName: this.excelTitles[index].editColumnName,
        edit: false,
        checked: this.excelTitles[index].checked,
        editExcelTitle: this.excelTitles[index].editExcelTitle,
        editColumnName: this.excelTitles[index].editColumnName
      }
    },
    editExcelTitle (index) {
      this.excelTitles[index].edit = true
    },
    deleteExcelTitle (index) {
      this.excelTitles.splice(index, 1)
    },
    cancelExcelTitle (index) {
      this.excelTitles[index] = {
        excelTitle: this.excelTitles[index].excelTitle,
        columnName: this.excelTitles[index].columnName,
        edit: false,
        checked: this.excelTitles[index].checked,
        editExcelTitle: this.excelTitles[index].excelTitle,
        editColumnName: this.excelTitles[index].columnName
      }
    },
    checkAll () {
      this.checkedAll = !this.checkedAll
      this.excelTitles.forEach(element => {
        element.checked = this.checkedAll
      })
    },
    checkItem (index) {
      this.excelTitles[index].checked = !this.excelTitles[index].checked
      let count = 0
      this.excelTitles.forEach(element => {
        if (element.checked) count++
      })
      if (count === this.excelTitles.length) {
        this.checkedAll = true
      } else {
        this.checkedAll = false
      }
    },
    clear () {
      this.excelTitles.length = 0
      this.fileName = ''
      this.fontFamily = ''
      this.fontSize = ''
      this.sheetName = ''
      this.exportCode = ''
      this.checkedAll = true
    }
  },
  persist: {
    enabled: true,
    strategies: [
      {
        key: 'import',
        storage: sessionStorage
      }
    ]
  }
})
