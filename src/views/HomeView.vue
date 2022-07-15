<script setup>
import ExportExcel from '@/components/ExportExcel.vue'
import ImportExcel from '@/components/ImportExcel.vue'
import { useExportStore } from '../stores/export'
import { ref } from 'vue'

const exportStore = useExportStore()
const navlink1 = ref(' active')
const navlink2 = ref('')
const tabpane1 = ref(' show active')
const tabpane2 = ref('')

if (exportStore.tab === 'import') {
  navlink1.value = ''
  navlink2.value = ' active'
  tabpane1.value = ''
  tabpane2.value = ' show active'
}
</script>

<script>
export default {
  components: {
    ExportExcel,
    ImportExcel
  }
}
</script>

<template lang="pug">
<main>
.container.mt-3
  .row.text-center
    h2.font-monospace ASP.NET NPOI Excel Export and Import
  .row.mt-2
    ul#pills-tab.nav.nav-pills.nav-justified.mb-3(role="tablist")
      li.nav-item(role="presentation")
        button#pills-export-tab.nav-link(:class="navlink1" data-bs-toggle="pill" data-bs-target="#pills-export" type="button" role="tab" aria-controls="pills-export" aria-selected="true" @click="exportStore.changeTab('export')") Export
      li.nav-item(role="presentation")
        button#pills-import-tab.nav-link(:class="navlink2" data-bs-toggle="pill" data-bs-target="#pills-import" type="button" role="tab" aria-controls="pills-import" aria-selected="false" @click="exportStore.changeTab('import')") Import
    #pills-tabContent.tab-content
      #pills-export.pt-2.tab-pane.fade(:class="tabpane1" role="tabpanel" aria-labelledby="pills-export-tab")
        <ExportExcel />
      #pills-import.pt-2.tab-pane.fade(:class="tabpane2" role="tabpanel" aria-labelledby="pills-import-tab")
        <ImportExcel />
</main>
</template>
