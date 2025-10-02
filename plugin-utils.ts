// plugin-utils.ts
// Утилиты для работы с плагином Pixso

/// <reference path="./pixso-types.d.ts" />

declare const pixso: any;

// Интерфейс для настроек экспорта
export interface ExportSettings {
  format: "SVG" | "PNG";
  includeBackgrounds: boolean;
  scale: number;
  svgOutlineText: boolean;
  svgSimplifyStroke: boolean;
}

// Настройки по умолчанию
export const defaultExportSettings: ExportSettings = {
  format: "SVG",
  includeBackgrounds: true,
  scale: 1,
  svgOutlineText: false,
  svgSimplifyStroke: true
};

// Функция для получения настроек экспорта
export function getExportSettings(): ExportSettings {
  try {
    const savedSettings = pixso.currentPage.getPluginData("exportSettings");
    if (savedSettings) {
      const parsed = JSON.parse(savedSettings);
      return { ...defaultExportSettings, ...parsed };
    }
  } catch (error) {
    console.warn("Failed to load export settings:", error);
  }
  return defaultExportSettings;
}

// Функция для сохранения настроек экспорта
export function saveExportSettings(settings: ExportSettings): void {
  try {
    pixso.currentPage.setPluginData("exportSettings", JSON.stringify(settings));
  } catch (error) {
    console.warn("Failed to save export settings:", error);
  }
}

// Функция для получения статистики последнего экспорта
export function getLastExportStats(): any {
  try {
    const stats = pixso.currentPage.getPluginData("lastExportStats");
    return stats ? JSON.parse(stats) : null;
  } catch (error) {
    console.warn("Failed to load export statistics:", error);
    return null;
  }
}

// Функция для получения информации о последней ошибке
export function getLastExportError(): any {
  try {
    const error = pixso.currentPage.getPluginData("lastExportError");
    return error ? JSON.parse(error) : null;
  } catch (error) {
    console.warn("Failed to load error information:", error);
    return null;
  }
}

// Функция для очистки данных плагина
export function clearPluginData(): void {
  try {
    const keys = pixso.currentPage.getPluginDataKeys();
    keys.forEach((key: string) => {
      if (key.startsWith("export") || key.startsWith("last")) {
        pixso.currentPage.setPluginData(key, "");
      }
    });
    console.log("Plugin data cleared");
  } catch (error) {
    console.warn("Failed to clear plugin data:", error);
  }
}

// Функция для валидации настроек экспорта
export function validateExportSettings(settings: any): ExportSettings {
  const validated: ExportSettings = { ...defaultExportSettings };
  
  if (settings.format === "SVG" || settings.format === "PNG") {
    validated.format = settings.format;
  }
  
  if (typeof settings.includeBackgrounds === "boolean") {
    validated.includeBackgrounds = settings.includeBackgrounds;
  }
  
  if (typeof settings.scale === "number" && settings.scale > 0 && settings.scale <= 10) {
    validated.scale = settings.scale;
  }
  
  if (typeof settings.svgOutlineText === "boolean") {
    validated.svgOutlineText = settings.svgOutlineText;
  }
  
  if (typeof settings.svgSimplifyStroke === "boolean") {
    validated.svgSimplifyStroke = settings.svgSimplifyStroke;
  }
  
  return validated;
}

// Функция для создания отчета об экспорте
export function createExportReport(stats: any, settings: ExportSettings): string {
  const report = `
Export Report
============
Date: ${new Date().toLocaleString()}
Total Frames: ${stats.totalFrames}
Successful: ${stats.processedFrames}
Failed: ${stats.failedFrames}
Success Rate: ${Math.round((stats.processedFrames / stats.totalFrames) * 100)}%

Settings:
- Format: ${settings.format}
- Include Backgrounds: ${settings.includeBackgrounds}
- Scale: ${settings.scale}
- SVG Outline Text: ${settings.svgOutlineText}
- SVG Simplify Stroke: ${settings.svgSimplifyStroke}
  `;
  
  return report.trim();
}

// Функция для экспорта отчета в файл (если поддерживается)
export function exportReport(report: string): void {
  try {
    // Сохраняем отчет как данные плагина
    pixso.currentPage.setPluginData("exportReport", report);
    console.log("Export report saved to plugin data");
  } catch (error) {
    console.warn("Failed to save export report:", error);
  }
}
