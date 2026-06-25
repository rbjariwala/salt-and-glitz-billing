import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// આ ફાઈલ વેરસેલને કન્ફર્મ કરશે કે આ એક React પ્રોજેક્ટ છે
export default defineConfig({
  plugins: [react()],
})
