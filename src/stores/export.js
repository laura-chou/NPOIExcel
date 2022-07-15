import { defineStore } from 'pinia'
export const useExportStore = defineStore({
  id: 'export',
  state: () => ({
    tab: 'export',
    excelTitles: [],
    fileName: '',
    fontFamily: '',
    fontSize: '',
    sheetName: '',
    fileExtension: '',
    exportCode: '',
    checkedAll: true
  }),
  getters: {
    copyData: (state) => state.excelTitles
  },
  actions: {
    changeTab (value) {
      this.tab = value
    },
    addExcelTitle (value) {
      this.excelTitles.push(value)
    },
    updateExcelTitle (index) {
      this.excelTitles[index] = {
        excelTitle: this.excelTitles[index].editExcelTitle,
        columnName: this.excelTitles[index].editColumnName,
        columnType: this.excelTitles[index].editColumnType,
        edit: false,
        checked: this.excelTitles[index].checked,
        editExcelTitle: this.excelTitles[index].editExcelTitle,
        editColumnName: this.excelTitles[index].editColumnName,
        editColumnType: this.excelTitles[index].editColumnType
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
        columnType: this.excelTitles[index].columnType,
        edit: false,
        checked: this.excelTitles[index].checked,
        editExcelTitle: this.excelTitles[index].excelTitle,
        editColumnName: this.excelTitles[index].columnName,
        editColumnType: this.excelTitles[index].columnType
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
    }
  },
  persist: {
    enabled: true,
    strategies: [
      {
        key: 'export',
        storage: sessionStorage
      }
    ]
  }
})
