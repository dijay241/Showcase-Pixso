// code.ts
// TS таргет: ES2019. Собираем из выделенных фреймов плоский список объектов
// (TEXT / SVG / IMAGE) с абсолютными координатами, размерами и поворотом.

// Импортируем типизацию Pixso API
/// <reference path="./pixso-types.d.ts" />

// Глобальное объявление pixso для TypeScript
declare const pixso: any;

  type ExportTextRun = {
    text: string;
    options: {
      fontFace?: string;
      fontSize?: number;
      color?: string;        // hex без '#'
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strike?: "sngStrike" | "dblStrike";
      charSpacing?: number;  // в пунктах
      lineSpacing?: number;  // в пунктах
      breakLine?: boolean;
      bullet?: boolean | { type?: string; code?: string; style?: string };
      indentLevel?: number;
    };
  };
  
  type ExportItemBase = {
    id: string;
    type: "TEXT" | "SVG" | "IMAGE" | "RECT" | "LINE" | "GROUP";
    xIn?: number; // inches
    yIn?: number;
    wIn?: number;
    hIn?: number;
    rotate?: number; // degrees
    name?: string;
    zIndex: number;
  };
  
  type ExportTextItem = ExportItemBase & {
    type: "TEXT";
    runs: ExportTextRun[];
    valign?: "top" | "middle" | "bottom";
    align?: "left" | "center" | "right";
    boxWIn: number;
    boxHIn: number;
    textwIn?: number;
  };
  
  type ExportSvgItem = ExportItemBase & {
    type: "SVG" | "LINE";
    svgBase64: string; // "image/svg+xml;base64,...."
    strokeWeight?: number;
  };
  
  type ExportImageItem = ExportItemBase & {
    type: "IMAGE";
    mime: "image/png" | "image/jpeg";
    dataBase64: string; // "image/png;base64,...."
  };

  type ExportRectItem = ExportItemBase & {
    type: "RECT" | "SVG";
    fillHex: string;
    borderHex?: string;
    borderWidth?: number;
    borderStyle?: string;
    cornerRadius?: number;
    topLeftRadius?: number;
    topRightRadius?: number;
    bottomLeftRadius?: number;
    bottomRightRadius?: number;
  };

  type ExportLineItem = ExportItemBase & {
    type: "LINE";
    strokeHex: string;
    strokePt: number; // thickness in points
    strokeOpacity: number; // opacity 0-1
    cap?: "round" | "square" | "flat";
    dash?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot";
    beginArrow?: string;
    opacity?: number;
    endArrow?: string;
    // PptxGenJS specific properties
    lineWidth?: number; // width in inches
    lineHeight?: number; // height in inches (for vertical lines)
    startX?: number; // start X coordinate in inches
    startY?: number; // start Y coordinate in inches
    endX?: number; // end X coordinate in inches
    endY?: number; // end Y coordinate in inches
  };

  type ExportGroupItem = ExportItemBase & {
    type: "GROUP";
    children: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem>;
  };
  
  type ExportFrame = {
    id: string;
    name: string;
    widthPx: number;
    heightPx: number;
    slideWIn: number;  // = widthPx / 96
    slideHIn: number;  // = heightPx / 96
    bgColorHex?: string;
    items: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem>;
  };
  
  // Интерфейс для настроек экспорта
  interface ExportSettings {
    format: "SVG" | "PNG";
    includeBackgrounds: boolean;
    scale: number;
    svgOutlineText: boolean;
    svgSimplifyStroke: boolean;
  }

  // Настройки по умолчанию
  const defaultExportSettings: ExportSettings = {
    format: "SVG",
    includeBackgrounds: true,
    scale: 1,
    svgOutlineText: false,
    svgSimplifyStroke: true
  };

  // Функция для получения настроек экспорта
  function getExportSettings(): ExportSettings {
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
  function saveExportSettings(settings: ExportSettings): void {
    try {
      pixso.currentPage.setPluginData("exportSettings", JSON.stringify(settings));
    } catch (error) {
      console.warn("Failed to save export settings:", error);
    }
  }

  // Функция для логирования ошибок экспорта
  function logExportError(node: any, error: any, format: string): void {
    console.error(`Export failed for ${node.name} (${node.type}) as ${format}:`, {
      error: error.message || error,
      nodeId: node.id,
      nodeSize: { 
        width: (node as any).width || 0, 
        height: (node as any).height || 0 
      },
      timestamp: new Date().toISOString()
    });
  }

  // Функция для безопасного экспорта узла
  async function safeExportNode(node: any, settings: ExportSettings): Promise<Uint8Array | null> {
    try {
      const exportSettings: any = {
        format: settings.format,
        constraint: { type: "SCALE", value: settings.scale },
        svgOutlineText: settings.svgOutlineText,
        svgIdAttribute: false,
        svgSimplifyStroke: settings.svgSimplifyStroke
      };

      return await node.exportAsync(exportSettings);
    } catch (primaryError) {
      logExportError(node, primaryError, settings.format);
      
      // Fallback: пробуем PNG если SVG не сработал
      if (settings.format === "SVG") {
        try {
          console.warn(`Falling back to PNG export for ${node.name}`);
          return await node.exportAsync({
            format: "PNG",
            constraint: { type: "SCALE", value: settings.scale }
          });
        } catch (fallbackError) {
          logExportError(node, fallbackError, "PNG");
          return null;
        }
      }
      
      return null;
    }
  }

  pixso.on("run", async ({ command }: { command: string }) => {

    // Получаем выбранные фреймы или все фреймы на странице
    const selection = pixso.currentPage.selection.length ? pixso.currentPage.selection : pixso.currentPage.children;
    const frames = selection.filter((n: any) =>
      n.type === "FRAME" || n.type === "COMPONENT" || n.type === "INSTANCE"
    ) as any[];
  
    // Если нет фреймов, закрываем плагин
    if (!frames.length) {
      pixso.closePlugin("There are no frames to export");
      return;
    }

    // Загружаем настройки экспорта
    const exportSettings = getExportSettings();
  
    // Показываем сообщение сразу при начале обработки
    let loadingMessage: any = null;
    loadingMessage = pixso.loading("Exporting frames...");

    const scratch = pixso.createFrame?.() ?? pixso.createRectangle(); // на случай, если frame недоступен
    scratch.name = '__tmp_flatten_scratch__';
    (scratch as any).resize(1400, 800);
    (scratch as any).visible = false;
    (scratch as any).locked = true;
    (scratch as any).x = 0;
    (scratch as any).y = 0;
    
    try {

      const exportFrames: ExportFrame[] = [];
      let zCounter = 0;
      let processedFrames = 0;
      let failedFrames = 0;
  
      
      // Обрабатываем фреймы последовательно для лучшей стабильности
      for (let i = 0; i < frames.length; i++) {

        const frame = frames[i];
        const masterFrameId = frame.id;

        try {
      
          const { x: frameAbsX, y: frameAbsY } = absOrigin(frame);
          const slideWIn = pxToIn(frame.width);
          const slideHIn = pxToIn(frame.height);
      
          // Получаем фон фрейма согласно Pixso FrameNode API
          const bgHex = frameBackgroundHex(frame);
  
          const items: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem> = [];
      
          // Получаем элементы из фрейма: сначала фреймы, затем дети, затем прочие элементы
          const walk = async (node: any, parentGroup?: ExportGroupItem) => {

            // Если элемент невидим, пропускаем его
            if (!node.visible) return;
        
            // Получаем абсолютные координаты элемента согласно Pixso API
            const { x: absX, y: absY } = absOrigin(node);
            // Преобразуем координаты в дюймы
            const xIn = pxToIn(absX - frameAbsX);
            const yIn = pxToIn(absY - frameAbsY);

            // Получаем поворот согласно Pixso API
            let rotate = 0;
            if ('rotation' in node && typeof (node as any).rotation === 'number') {
              // Используем свойство rotation (диапазон [-180, 180])
              rotate = (node as any).rotation;
            } else if (node.relativeTransform) {
              // Вычисляем из relativeTransform согласно документации
              rotate = rotationFromTransform(node.relativeTransform);
            } else if (node.absoluteTransform) {
              // Fallback: вычисляем из absoluteTransform
              rotate = rotationFromTransform(node.absoluteTransform);
            }
        
            // Нормализуем угол поворота
            while (rotate > 180) rotate -= 360;
            while (rotate < -180) rotate += 360;
        
            // Получаем размеры элемента согласно Pixso API
            const wIn = hasSize(node) ? pxToIn(node.width) : undefined;
            const hIn = hasSize(node) ? pxToIn(node.height) : undefined;

            if ("children" in node && (node as any).children) {
              // Для фреймов и других контейнеров просто рекурсивно обходим детей
              for (const child of (node as any).children as any[]) {
                await walk(child, parentGroup);
              }
            }

            // Обработка текстовых элементов
            if (node.type === "TEXT") {

              const textNode = node;
              const runs: ExportTextRun[] = [];
  
              // В Pixso API используем методы getRange* для получения свойств текста
              // Создаем один сегмент для всего текста, так как в Pixso нет getStyledTextSegments
              const textLength = textNode.characters.length;
              const segments = [{
                start: 0,
                end: textLength,
                fontName: textNode.getRangeFontName ? textNode.getRangeFontName(0, textLength) : (textNode as any).fontName,
                fontSize: textNode.getRangeFontSize ? textNode.getRangeFontSize(0, textLength) : (textNode as any).fontSize,
                fills: (textNode as any).fills,
                textDecoration: textNode.getRangeTextDecoration ? textNode.getRangeTextDecoration(0, textLength) : (textNode as any).textDecoration,
                letterSpacing: textNode.getRangeLetterSpacing ? textNode.getRangeLetterSpacing(0, textLength) : (textNode as any).letterSpacing,
                lineHeight: textNode.getRangeLineHeight ? textNode.getRangeLineHeight(0, textLength) : (textNode as any).lineHeight
              }];
          
              // В Pixso API нет getRangeListOptions, пропускаем обработку списков
              for (const seg of segments) {

                const str = textNode.characters.substring(seg.start, seg.end);
                if (!str.length) continue;

                const fill = firstSolidPaint(seg.fills);
                const colorHex = fill ? rgbToHex(fill.color, fill.opacity) : "000000";
            
                // Обработка textDecoration согласно Pixso API
                const underline = seg.textDecoration === "UNDERLINE";
                const strike = seg.textDecoration === "STRIKETHROUGH" ? "sngStrike" : undefined;
            
                // Обработка шрифтов согласно Pixso API
                let fontFamily = "Arial";
                let fontSize = 12;
                let isBold = false;
                let isItalic = false;
            
                // Обработка FontName согласно документации Pixso API
                if (seg.fontName && typeof seg.fontName === 'object') {
                  if (seg.fontName.family) {
                    fontFamily = seg.fontName.family;
                  }
                  if (seg.fontName.style) {
                    isBold = /(bold|semibold|demi|medium|black|heavy)/i.test(seg.fontName.style);
                    isItalic = /(italic|oblique)/i.test(seg.fontName.style);
                  }
                }
            
                // Обработка fontSize согласно документации Pixso API
                if (seg.fontSize && typeof seg.fontSize === 'number') {
                  fontSize = seg.fontSize;
                }

                // Обработка LetterSpacing согласно Pixso API
                let charSpacingPt: number | undefined;
                if (seg.letterSpacing && typeof seg.letterSpacing === 'object') {
                  const letterSpacing = seg.letterSpacing as any;
                  if (letterSpacing.unit === "PIXELS") {
                    charSpacingPt = pxToPt(letterSpacing.value);
                  } else if (letterSpacing.unit === "PERCENT") {
                    charSpacingPt = pxToPt(fontSize * letterSpacing.value / 100);
                  } else if (letterSpacing.unit === "PT") {
                    charSpacingPt = letterSpacing.value;
                  }
                }

                // Обработка LineHeight согласно Pixso API
                let lineSpacing: number | undefined;
                if (seg.lineHeight && typeof seg.lineHeight === 'object') {
                  const lineHeight = seg.lineHeight as any;
                  console.log("lineHeight", lineHeight);
                  if (lineHeight.unit === "PIXELS") {
                    lineSpacing = pxToPt(lineHeight.value);
                  } else if (lineHeight.unit === "PERCENT") {
                    lineSpacing = pxToPt(fontSize * lineHeight.value / 100);
                  } else if (lineHeight.unit === "PT") {
                    lineSpacing = lineHeight.value;
                  } else if (lineHeight.unit === "AUTO") {
                    lineSpacing = pxToPt(fontSize * 1.25);
                  }
                } else {
                  // Если lineHeight не задан, используем стандартный 120%
                  lineSpacing = pxToPt(fontSize * 1.25);
                }

                // В Pixso API нет поддержки списков, пропускаем
                let bullet: boolean | { type?: string; code?: string; style?: string } | undefined;

                const runOptions: any = {
                  fontFace: fontFamily,
                  fontSize: Math.round(pxToPt(fontSize)),
                  color: colorHex,
                  underline: underline || undefined,
                  strike: (strike === "sngStrike" || strike === "dblStrike") ? strike : undefined,
                  bold: isBold || undefined,
                  italic: isItalic || undefined,
                  charSpacing: charSpacingPt,
                  lineSpacing: lineSpacing,
                  bullet: bullet
                };
            
            
                runs.push({
                  text: str,
                  options: runOptions
                });
              }
  
              const align = mapAlign(textNode.textAlignHorizontal);
              const valign = mapVAlign(textNode.textAlignVertical);
              let textwIn = wIn ? wIn : 0;
              const wDelta = textwIn * 0.15;
  

              // Проверяем, что у нас есть хотя бы один run с текстом
              if (runs.length === 0) {
                // Если нет runs, создаем один с базовыми настройками из самого текстового узла
                const fallbackFill = firstSolidPaint((textNode as any).fills);
                const fallbackColor = fallbackFill ? rgbToHex(fallbackFill.color, fallbackFill.opacity) : "000000";
                
                // Используем методы Pixso API для получения свойств
                const fallbackFontSize = textNode.getRangeFontSize ? textNode.getRangeFontSize(0, textNode.characters.length) : 12;
                const fallbackFontName = textNode.getRangeFontName ? textNode.getRangeFontName(0, textNode.characters.length) : { family: "Arial", style: "Regular" };
                
                runs.push({
                  text: textNode.characters || "",
                  options: {
                    fontFace: fallbackFontName.family || "Arial",
                    fontSize: Math.round(pxToPt(fallbackFontSize)),
                    color: fallbackColor,
                    bold: /(bold|semibold|demi|medium|black|heavy)/i.test(fallbackFontName.style || '') || undefined,
                    italic: /(italic|oblique)/i.test(fallbackFontName.style || '') || undefined
                  }
                });
              }

              // Корректируем координаты для повернутого текста
              let adjustedXIn = xIn;
              let adjustedYIn = yIn;
          
              // Применяем корректировку координат для повернутого текста
              if (Math.abs(rotate) > 0.01) {
                const radians = (rotate * Math.PI) / 180;
                const textWidth = textwIn || 1;
                const textHeight = hIn || 1;
                
                // Корректировка для разных углов поворота
                if (Math.abs(rotate - 90) < 0.01) {
                  // Поворот на 90°: смещение по Y на ширину текста
                  adjustedYIn = yIn + textWidth;
                } else if (Math.abs(rotate + 90) < 0.01) {
                  // Поворот на -90°: смещение по X на высоту текста
                  adjustedXIn = xIn + textHeight;
                } else if (Math.abs(rotate - 180) < 0.01 || Math.abs(rotate + 180) < 0.01) {
                  // Поворот на 180°: смещение по обеим осям
                  adjustedXIn = xIn + textWidth;
                  adjustedYIn = yIn + textHeight;
                } else {
                  // Для других углов применяем общую формулу
                  const cos = Math.cos(radians);
                  const sin = Math.sin(radians);
                  
                  // Корректировка с учетом центра поворота
                  const centerX = textWidth / 2;
                  const centerY = textHeight / 2;
                  
                  adjustedXIn = xIn + centerX - (centerX * cos - centerY * sin);
                  adjustedYIn = yIn + centerY - (centerX * sin + centerY * cos);
                }
              }

              textwIn = textwIn + wDelta;

              if(align === "center") {
                adjustedXIn = adjustedXIn - wDelta / 2;
              } else if(align === "right") {
                adjustedXIn = adjustedXIn - wDelta;
              }

              const textItem: ExportTextItem = {
                id: textNode.id,
                type: "TEXT",
                xIn: adjustedXIn, yIn: adjustedYIn,
                textwIn, hIn,
                boxWIn: textwIn || 1,
                boxHIn: hIn || 1,
                rotate: -rotate, // Инвертируем знак угла для PowerPoint
                runs,
                align, valign,
                name: textNode.name,
                zIndex: zCounter++
              };

              items.push(textItem);

            }

            else if (node.type === "RECTANGLE" || node.type === "ELLIPSE" || node.type === "POLYGON" || node.type === "STAR") {

              // Проверяем объект на нужный тип и размер
              if (!hasRenderableSize(node)) return;
              
              const parentNode = node.parent;
              let nodeToExport = node;
              
              // Создаем копию узла для flatten, чтобы не изменять оригинал
              const clonedNode = (node as any).clone();
              clonedNode.x = node.x;
              clonedNode.y = node.y;

              if(masterFrameId !== parentNode.id) {
                const absolutePosition = calculateAbsolutePosition(node, clonedNode, masterFrameId);
                clonedNode.x = absolutePosition.x;
                clonedNode.y = absolutePosition.y;
              }

              // Применяем flatten к повернутому объекту
              nodeToExport = pixso.flatten([clonedNode], scratch);

              const data = await safeExportNode(nodeToExport, exportSettings);
              
              if (data) {
                const svgBase64 = "data:image/svg+xml;base64," + pixso.base64Encode(data);
                const svgItem: ExportSvgItem = {
                  id: node.id,
                  type: "SVG",
                  xIn: pxToIn(nodeToExport.x), 
                  yIn: pxToIn(nodeToExport.y), 
                  wIn: pxToIn(nodeToExport.width), 
                  hIn: pxToIn(nodeToExport.height),
                  svgBase64,
                  name: node.name,
                  zIndex: zCounter++
                };

                items.push(svgItem);

              }
              
            }

            else if (node.type === "LINE") {
              lineProcessing(node, rotate, zCounter, items);
            }

            else if (isVectorNode(node)) {
              // Специальная обработка для линий (в Pixso линии имеют тип VECTOR)
              const isLine = node.width < 1 || node.height < 1;
              
              if (isLine) {
                lineProcessing(node, rotate, zCounter, items);
                return; // Выходим из функции для линий
              }
              
              if (!hasRenderableSize(node)) return;
              
              // Для векторных объектов используем нативный поворот PowerPoint для стандартных углов
              let nodeToExport = node;
              let useFlattened = false;

              
              // Применяем flatten только для нестандартных углов поворота
              const isStandardRotation = Math.abs(rotate) < 0.01 || 
                                       Math.abs(rotate - 90) < 0.01 || 
                                       Math.abs(rotate + 90) < 0.01 || 
                                       Math.abs(rotate - 180) < 0.01 || 
                                       Math.abs(rotate + 180) < 0.01;
              
              if (!isStandardRotation && Math.abs(rotate) > 0.01) {
                try {
                  // Создаем копию узла для flatten, чтобы не изменять оригинал
                  const clonedNode = (node as any).clone();
                  // Применяем flatten к повернутому узлу
                  nodeToExport = pixso.flatten([clonedNode]);
                  useFlattened = true;
                } catch (flattenError) {
                  console.warn(`Failed to flatten ${node.name}:`, flattenError);
                  nodeToExport = node;
                  useFlattened = false;
                }
              }

              // Используем безопасный экспорт
              const data = await safeExportNode(nodeToExport, exportSettings);
              
              if (data) {
                const svgBase64 = "data:image/svg+xml;base64," + pixso.base64Encode(data);
                const strokeWeight = (nodeToExport as any).strokeWeight;

                const svgItem: ExportSvgItem = {
                  id: node.id,
                  type: "SVG",
                  // Используем координаты и размеры оригинального узла, а не сплющенного
                  xIn, yIn, wIn, hIn: hIn || (strokeWeight ? strokeWeight * 0.01 : 0.01),
                  rotate: useFlattened ? 0 : rotate, // После flatten поворот уже применен к геометрии
                  svgBase64,
                  name: node.name,
                  zIndex: zCounter++
                };


                if (parentGroup) {
                  parentGroup.children.push(svgItem);
                } else {
                  items.push(svgItem);
                }
              } else {
                // Если векторный узел не может быть экспортирован, создаем прямоугольник как fallback
                console.warn(`Creating fallback rectangle for vector node ${node.name}`);
                const fills = (node as any).fills;
                const fill = firstSolidPaint(fills);
                const fillHex = fill ? rgbToHex(fill.color, fill.opacity) : "000000";

                const rectItem: ExportRectItem = {
                  id: node.id + "::fallback",
                  type: "RECT",
                  xIn, yIn, wIn, hIn,
                  rotate,
                  fillHex,
                  name: node.name + " (fallback)",
                  zIndex: zCounter++
                };

                if (parentGroup) {
                  parentGroup.children.push(rectItem);
                } else {
                  items.push(rectItem);
                }
              }
            }
        else if (hasImageFill(node)) {
          if (!hasRenderableSize(node)) return;
          
          
          // Для изображений используем PNG формат
          const imageSettings: ExportSettings = {
            ...exportSettings,
            format: "PNG"
          };
          
          const bytes = await safeExportNode(node, imageSettings);
          
          if (bytes) {
            const dataBase64 = "data:image/png;base64," + pixso.base64Encode(bytes);
            
            const imageItem: ExportImageItem = {
              id: node.id,
              type: "IMAGE",
              xIn, yIn, wIn, hIn,
              rotate,
              mime: "image/png",
              dataBase64,
              name: node.name,
              zIndex: zCounter++
            };
            
            if (parentGroup) {
              parentGroup.children.push(imageItem);
            } else {
              items.push(imageItem);
            }
            
          } else {
            console.warn(`Failed to export image node ${node.name}, skipping`);
          }
        }
      };
      await walk(frame);
  
          exportFrames.push({
            id: frame.id,
            name: frame.name,
            widthPx: frame.width,
            heightPx: frame.height,
            slideWIn, slideHIn,
            bgColorHex: bgHex || undefined,
            items
          });
          
          processedFrames++;
          
        } catch (frameError) {
          failedFrames++;
          console.error(`Failed to process frame ${frame.name}:`, frameError);
          
          // Добавляем пустой фрейм для сохранения структуры
          exportFrames.push({
            id: frame.id,
            name: frame.name + " (failed)",
            widthPx: frame.width,
            heightPx: frame.height,
            slideWIn: pxToIn(frame.width),
            slideHIn: pxToIn(frame.height),
            bgColorHex: undefined,
            items: []
          });
        }
      }
      
      
      // Сохраняем статистику экспорта
      const exportStats = {
        totalFrames: frames.length,
        processedFrames,
        failedFrames,
        timestamp: new Date().toISOString(),
        settings: exportSettings
      };
      
      try {
        pixso.currentPage.setPluginData("lastExportStats", JSON.stringify(exportStats));
      } catch (error) {
        console.warn("Failed to save export statistics:", error);
      }
    
    // Скрываем сообщение о загрузке после обработки всех фреймов
    if (loadingMessage) {
      loadingMessage.cancel();
    }
  
    pixso.ui.onmessage = (msg: any) => {
      if (msg.type === "done") {
        pixso.closePlugin("Export complete");
      } else if (msg.type === "close") {
        // Закрываем плагин по запросу от UI
        pixso.closePlugin();
      }
    };

      // Без UI: отображение выключено, сразу отправляем данные для авто-экспорта
      pixso.showUI((globalThis as any).__html__, { visible: false });
      const fileNameBase = exportFrames.length === 1 ? exportFrames[0].name : pixso.currentPage.name;

      const serializedData = JSON.parse(JSON.stringify({ 
        type: "AUTO", 
        frames: exportFrames, 
        fileNameBase 
      }));
      pixso.ui.postMessage(serializedData);
      
    } catch (error) {
      // Скрываем сообщение о загрузке в случае ошибки
      if (loadingMessage) {
        loadingMessage.cancel();
      }
      
      // Логируем детальную информацию об ошибке
      console.error("Export process failed:", {
        error: (error as any)?.message || error,
        stack: (error as any)?.stack,
        timestamp: new Date().toISOString(),
        framesCount: frames.length
      });
      
      // Сохраняем информацию об ошибке
      try {
        const errorInfo = {
          message: (error as any)?.message || "Unknown error",
          timestamp: new Date().toISOString(),
          framesCount: frames.length
        };
        pixso.currentPage.setPluginData("lastExportError", JSON.stringify(errorInfo));
      } catch (saveError) {
        console.warn("Failed to save error information:", saveError);
      }
      
      if ((error as any)?.message !== "Export cancelled") {
        const errorMessage = `Export failed: ${(error as any)?.message || "Unknown error"}`;
        pixso.notify(errorMessage, { timeout: 5000, error: true });
      }
      pixso.closePlugin();
    }

    (scratch as any).remove?.();

  });
  
  // ==== helpers ====
  
  // async function traverse(node: any, fn: (n: any) => void | Promise<void>) : Promise<void> {
  //   // Сохраняем порядок слоёв: от заднего к переднему
  //   if ("children" in node) {
  //     for (const child of node.children as any[]) {
  //       await traverse(child, fn);
  //     }
  //   }
  //   // Пропускаем сам фрейм как геом. элемент, но обрабатываем содержимое
  //   if (node.type !== "FRAME" && node.type !== "COMPONENT" && node.type !== "INSTANCE") {
  //     await fn(node);
  //   }
  // }
  
  function absOrigin(n: any) {
    // absoluteTransform = [[a,b,tx],[c,d,ty]]
    const m = (n as any).absoluteTransform as any;
    return { x: m[0][2], y: m[1][2] };
  }
  
  function rotationFromTransform(m: any): number {
    const a = m[0][0], b = m[0][1], c = m[1][0], d = m[1][1];
    
    // Вычисляем угол поворота из матрицы трансформации согласно Pixso API
    // Формула: Math.atan2(-relativeTransform[1][0], relativeTransform[0][0])
    let angle = Math.atan2(-c, a) * 180 / Math.PI;
    
    // Нормализуем угол в диапазон -180 до 180 (согласно документации)
    while (angle > 180) angle -= 360;
    while (angle < -180) angle += 360;
    
    return Math.round(angle * 1000) / 1000;
  }
  
  function firstSolidPaint(paints: any) {
    if (!Array.isArray(paints)) return null;
    const p = paints.find((p: any) => p.type === "SOLID") as any;
    return p || null;
  }

  function frameBackgroundHex(n: any) {
    // Согласно Pixso FrameNode API, используем fills вместо backgrounds
    const paints = (n as any).fills as any;
    const p = firstSolidPaint(paints);
    if (!p) return null;
    return rgbToHex(p.color, (p.opacity === undefined || p.opacity === null) ? 1 : p.opacity);
  }
  
  function rgbToHex(c: any, opacity: number = 1) {
    const r = Math.round(c.r * 255);
    const g = Math.round(c.g * 255);
    const b = Math.round(c.b * 255);
    // Игнорируем прозрачность для текста — PowerPoint не поддерживает прозрачные глифы напрямую
    return [r, g, b].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  
  function mapAlign(a: "LEFT" | "CENTER" | "RIGHT" | "JUSTIFIED"): "left" | "center" | "right" {
    return a === "CENTER" ? "center" : a === "RIGHT" ? "right" : "left";
  }
  
  function mapVAlign(v: "TOP" | "CENTER" | "BOTTOM"): "top" | "middle" | "bottom" {
    return v === "CENTER" ? "middle" : v === "BOTTOM" ? "bottom" : "top";
  }
  
  function isFont(f: any): f is any {
    return !!f && typeof f === "object" && "family" in f;
  }

  // Функция для суммирования координат по всем parent нодам до masterFrameId
  function calculateAbsolutePosition(node: any, flattenedNode: any, masterFrameId: string): any {
    let totalX = flattenedNode.x;
    let totalY = flattenedNode.y;
    let currentParent = node.parent;
    // Проходим по всем parent нодам до masterFrameId
    while (currentParent && currentParent.id !== masterFrameId) {
      if(currentParent.type === "FRAME") {
        totalX += currentParent.x || 0;
        totalY += currentParent.y || 0;
      }
      currentParent = currentParent.parent;
    }
    return { x: totalX, y: totalY };
  };
  
  function isVectorNode(n: any): boolean {
    return (
      (n.type === "VECTOR" ||
        n.type === "BOOLEAN_OPERATION") &&
      "exportAsync" in n
    );
  }

  // // Функция для определения линии в Pixso
  // function isLineNode(n: any): boolean {
  //   // В Pixso линии имеют тип VECTOR, но у них width < 1 или height < 1
  //   return n.type === "VECTOR" && (n.width < 1 || n.height < 1);
  // }

  // Функция для анализа vectorNetwork для определения стрелок
  function analyzeVectorNetworkForArrows(node: any): { beginArrow: string; endArrow: string } {
    
    let beginArrow = "none";
    let endArrow = "none";
    
    try {
      if (node.vectorNetwork) {
        const vectorNetwork = node.vectorNetwork;
        
        // Проверяем vertices
        if (vectorNetwork.vertices && Array.isArray(vectorNetwork.vertices)) {
          
          if (vectorNetwork.vertices.length >= 2) {
            const startVertex = vectorNetwork.vertices[0];
            const endVertex = vectorNetwork.vertices[vectorNetwork.vertices.length - 1];
            
            if (startVertex && startVertex.strokeCap) {
              beginArrow = mapLineEnding(endVertex.strokeCap);
            }
            
            if (endVertex && endVertex.strokeCap) {
              endArrow = mapLineEnding(startVertex.strokeCap);
            }
          }
        }
        
        // Fallback: проверяем segments
        if (vectorNetwork.segments && Array.isArray(vectorNetwork.segments)) {
          
          for (const segment of vectorNetwork.segments) {
            
            if (segment.startCap && segment.startCap !== "NONE") {
              beginArrow = mapLineEnding(segment.startCap);
            }
            
            if (segment.endCap && segment.endCap !== "NONE") {
              endArrow = mapLineEnding(segment.endCap);
            }
          }
        }
      } 
      
      // Fallback: проверяем прямые свойства стрелок
      if (node.strokeStartArrow !== undefined) {
        beginArrow = mapLineEnding(node.strokeStartArrow);
      }
      
      if (node.strokeEndArrow !== undefined) {
        endArrow = mapLineEnding(node.strokeEndArrow);
      }
      
    } catch (error) {
      console.error("Error analyzing vectorNetwork:", error);
      console.error("Error details:", (error as any)?.message);
      console.error("Node data:", node);
    }
    
    const result = { beginArrow, endArrow };
    
    return result;
  }

  // Функция для маппинга dashPattern в PptxGenJS формат
  function mapDashPattern(dashPattern: number[] | undefined): "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" {
    if (!dashPattern || dashPattern.length === 0) {
      return "solid";
    }
    
    // Анализируем паттерн штриховки согласно документации Pixso
    if (dashPattern.length === 2) {
      const [dash, gap] = dashPattern;
      if (dash > 0 && gap > 0) {
        return "dash";
      } else if (dash > 0 && gap === 0) {
        return "dot";
      }
    } else if (dashPattern.length === 4) {
      const [dash1, gap1, dash2, gap2] = dashPattern;
      if (dash1 > 0 && gap1 > 0 && dash2 > 0 && gap2 > 0) {
        return "dashDot";
      }
    }
    
    // Для сложных паттернов возвращаем dash как fallback
    return "dash";
  }

  function hasImageFill(n: any): boolean {
    // Считаем растровыми только узлы с заливкой типа IMAGE
    const anyNode = n as any;
    const fills = anyNode && anyNode.fills;
    if (!fills || !Array.isArray(fills)) return false;
    return fills.some((p: any) => p && p.type === 'IMAGE');
  }

  function hasSize(n: any): n is any & { width: number; height: number } {
    return "width" in n && "height" in n;
  }

  function hasRenderableSize(n: any): boolean {
    if (!("exportAsync" in n)) return false;
    if (!hasSize(n)) return false;
    // Разрешаем линии (w == 0 или h == 0), но отсекаем полностью нулевые
    if (n.width === 0 && n.height === 0) return false;
    // Контейнеры не рендерим как отдельные картинки
    if (n.type === "GROUP" || n.type === "FRAME" || n.type === "COMPONENT" || n.type === "INSTANCE") {
      const g = n as any;
      if (!('children' in g) || !g.children || g.children.length === 0) return false;
      return false;
    }
    return true;
  }
  
  function pxToIn(px: number | undefined):number { return px ? px / 96 : 0; }       // 96 px = 1 in
  
  function pxToPt(px: number | undefined):number { return px ? px * 0.75 : 0; }     // 1 px = 0.75 pt
  
  // function frameBorder(n: any) {
  //   const strokes = (n as any).strokes as any;
  //   const strokeWeight = (n as any).strokeWeight;
  //   const strokeAlign = (n as any).strokeAlign;
    
  //   if (!strokes || !Array.isArray(strokes) || !strokeWeight) return null;
    
  //   const stroke = firstSolidPaint(strokes);
  //   if (!stroke) return null;
    
  //   return {
  //     color: rgbToHex(stroke.color, (stroke.opacity === undefined || stroke.opacity === null) ? 1 : stroke.opacity),
  //     width: pxToPt(strokeWeight),
  //     style: strokeAlign === 'CENTER' ? 'solid' : 'solid' // можно расширить для разных стилей
  //   };
  // }

  function projectionOnXDegrees(length: number, angleDegrees: number): number {
    const EPS = 1e-8;
    const angleRadians = (angleDegrees * Math.PI) / 180;
    const value = length * Math.cos(angleRadians);
    return Math.abs(value) < EPS ? 0 : value;
  }

  function projectionOnYDegrees(length: number, angleDegrees: number): number {
    const EPS = 1e-8;
    const angleRadians = (angleDegrees * Math.PI) / 180;
    const value = length * Math.sin(angleRadians);
    return Math.abs(value) < EPS ? 0 : value;
  }

  function mapLineEnding(type: string): "arrow" | "diamond" | "oval" | "stealth" | "triangle" | "none" {
    const mapping: Record<string, "arrow" | "diamond" | "oval" | "stealth" | "triangle" | "none"> = {
      NONE: "none",
      ARROW_LINES: "arrow",
      ROUND: "oval",
      SQUARE: "diamond",
      ARROW_EQUILATERAL: "triangle",
      TRIANGLE_FILLED: "triangle",
      HOLLOW_ROUND: "oval",
      CIRCLE_FILLED: "oval",
      VERTICAL_LINE: "oval",
      DIAMOND_FILLED: "diamond"
    };
  
    return mapping[type] ?? "none"; // если не найдено, возвращаем "none"
  }

  function lineProcessing(
    node: any, 
    rotate: number, 
    zCounter: number,
    items: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem>) {

    try {

      // Анализируем vectorNetwork для определения стрелок
      const arrowInfo = analyzeVectorNetworkForArrows(node as any);
    
      // Экспортируем как линию
      const strokeWeight = (node as any).strokeWeight || 1;
      const stroke = firstSolidPaint((node as any).strokes);
      const strokeHex = stroke ? rgbToHex(stroke.color, stroke.opacity) : "000000";
    
      // Определяем направление линии
      const isHorizontal = node.width > node.height;
      const lineWidth = isHorizontal ? node.width : 0;  // Для горизонтальных = длина, для вертикальных = 0
      const lineHeight = isHorizontal ? 0 : node.height;  // Для горизонтальных = 0, для вертикальных = длина
      
      const projX = (pxToIn(lineWidth) - projectionOnXDegrees(pxToIn(lineWidth), rotate))/2;
      const projY = projectionOnYDegrees(pxToIn(lineWidth), rotate)/2;

      let dX = 0;
      let dY = 0;

      // Вычисляем смещение для повернутых линий
      if (Math.abs(rotate) === 180) {
        dX = projX;
        dY = 0;
      } else if (rotate > 0) {
        dX = projX;
        dY = projY;
      } else if (rotate < 0) {
        dX = projX;
        dY = -Math.abs(projY);
      }

      const lineItem: ExportLineItem = {
        id: node.id,
        type: "LINE",
        xIn: (pxToIn(node.x) - dX), 
        yIn: (pxToIn(node.y) - dY),
        wIn: pxToIn(lineWidth),
        hIn: pxToIn(lineHeight),
        rotate: -rotate,
        opacity: (node as any).opacity,
        strokeHex,
        strokePt: pxToPt(strokeWeight),
        strokeOpacity: stroke ? (stroke.opacity || 1) : 1,
        cap: (node as any).strokeCap || "flat",
        dash: mapDashPattern((node as any).dashPattern),
        beginArrow: arrowInfo.beginArrow,
        endArrow: arrowInfo.endArrow,
        name: node.name,
        zIndex: zCounter++
      };

      items.push(lineItem);
      
    } catch (error) {

      console.error("❌ Error processing LINE item:", error);
      console.error("Error details:", (error as any)?.message);
      console.error("Node data:", node);
      
    }
  }