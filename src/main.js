/* eslint-disable vue/multi-word-component-names */
import { createApp } from 'vue'
import { createPinia } from 'pinia'
import App from './App.vue'
import router from './router'
import VueFeather from 'vue-feather'
import 'bootstrap/dist/css/bootstrap.min.css'
import 'bootstrap'
import VueHighlightJS from 'vue3-highlightjs'
import 'highlight.js/styles/vs2015.css'
import './assets/style.stylus'
import Draggable from 'vuedraggable'
import piniaPersist from 'pinia-plugin-persist'
const pinia = createPinia()
pinia.use(piniaPersist)

const app = createApp(App)
app.component('draggable', Draggable)
app.component(VueFeather.name, VueFeather)
app.use(VueHighlightJS)
app.use(pinia)
app.use(router)
app.mount('#app')
