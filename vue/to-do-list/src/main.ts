import './assets/main.css'
import testBtn from './TodoFooter/testBtn.vue'

import { createApp } from 'vue'
import App from './App.vue'

const app = createApp(App)

app.mount('#app')

app.component('testBtn', testBtn)