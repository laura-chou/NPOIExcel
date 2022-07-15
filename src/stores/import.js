import { defineStore } from 'pinia'
import { useExportStore } from './export.js'
export const useImportStore = defineStore({
  id: 'import',
  state: () => ({
    fileName: '',
    filePath: '',
    saveFile: '',
    fileSizeLimit: '',
    importCode: ''
  }),
  getters: {
  },
  actions: {
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
