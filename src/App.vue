<script setup lang="ts">
import { ElConfigProvider } from 'element-plus'
import zhCn from 'element-plus/es/locale/lang/zh-cn'
import ReportGenerator from './components/ReportGenerator.vue'
</script>

<template>
  <el-config-provider :locale="zhCn">
    <ReportGenerator />
  </el-config-provider>
</template>

<style scoped>
</style>
