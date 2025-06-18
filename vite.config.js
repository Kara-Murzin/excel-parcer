import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Это имя вашего репозитория на GitHub.
// Оно должно совпадать с частью URL после вашего никнейма: https://ВАШ_НИКНЕЙМ.github.io/ИМЯ_ВАШЕГО_РЕПОЗИТОРИЯ/
const REPO_NAME = 'excel-parcer'; // <-- Убедитесь, что это ТОЧНОЕ имя вашего репозитория

export default defineConfig({
  plugins: [react()],
  base: `/${REPO_NAME}/`, // <-- ЭТО КРИТИЧЕСКИ ВАЖНО для GitHub Pages
});