<script>
import ExportExcel from '@/components/ExportExcel.vue'
import ImportExcel from '@/components/ImportExcel.vue'
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
  ul#pills-tab.nav.nav-pills.nav-justified.mb-3(role="tablist")
    li.nav-item(role="presentation")
      button#pills-export-tab.nav-link.active(data-bs-toggle="pill" data-bs-target="#pills-export" type="button" role="tab" aria-controls="pills-export" aria-selected="true") Export
    li.nav-item(role="presentation")
      button#pills-import-tab.nav-link(data-bs-toggle="pill" data-bs-target="#pills-import" type="button" role="tab" aria-controls="pills-import" aria-selected="false") Import
  #pills-tabContent.tab-content
    #pills-export.tab-pane.fade.show.active(role="tabpanel" aria-labelledby="pills-export-tab")
      <ExportExcel />
    #pills-import.tab-pane.fade(role="tabpanel" aria-labelledby="pills-import-tab") 尚未完成
      //- <ImportExcel />
</main>
</template>
