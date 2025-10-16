#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { v4 as uuidv4 } from 'uuid';
import { promises as fs } from 'fs';
import path from 'path';
import PptxGenJS from 'pptxgenjs';

// Export directory configuration
const EXPORT_DIR = '/tmp/protex-intelligence-file-exports';

/**
 * Calculate file size string from buffer
 */
function getFileSizeString(buffer) {
  const bytes = buffer.length;
  const kb = Math.ceil(bytes / 1024);
  return kb < 1024 ? `${kb} KB` : `${(kb / 1024).toFixed(2)} MB`;
}

/**
 * Ensure export directory exists, create if it doesn't
 */
async function ensureExportDirectory() {
  try {
    await fs.access(EXPORT_DIR);
    console.error(`✓ Export directory exists: ${EXPORT_DIR}`);
  } catch (error) {
    try {
      await fs.mkdir(EXPORT_DIR, { recursive: true });
      console.error(`✓ Created export directory: ${EXPORT_DIR}`);
    } catch (mkdirError) {
      console.error(`✗ Failed to create export directory: ${mkdirError.message}`);
      throw mkdirError;
    }
  }
}

/**
 * Write PPTX buffer to file system
 */
async function writePPTXToFile(pptxBuffer, filename) {
  await ensureExportDirectory();

  const filepath = path.join(EXPORT_DIR, filename);

  try {
    await fs.writeFile(filepath, pptxBuffer);
    console.error(`✓ File written: ${filepath}`);
    return filepath;
  } catch (error) {
    console.error(`✗ Failed to write file: ${error.message}`);
    throw error;
  }
}

// Create MCP server
const server = new Server(
  {
    name: 'pptx-export-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'pptx_export',
        description: 'Export data to PowerPoint (PPTX) format with full support for text, tables, charts, images, and shapes',
        inputSchema: {
          type: 'object',
          properties: {
            slides: {
              type: 'array',
              description: 'Array of slide objects containing elements to add',
              items: {
                type: 'object',
                properties: {
                  backgroundColor: {
                    type: 'string',
                    description: 'Slide background color (hex code without #)',
                  },
                  elements: {
                    type: 'array',
                    description: 'Array of elements to add to the slide (text, table, chart, image, shape)',
                    items: {
                      type: 'object',
                      properties: {
                        type: {
                          type: 'string',
                          enum: ['text', 'table', 'chart', 'image', 'shape'],
                          description: 'Type of element to add',
                        },
                        text: {
                          type: ['string', 'array'],
                          description: 'For type=text: String or array of text objects with formatting',
                        },
                        rows: {
                          type: 'array',
                          description: 'For type=table: Array of rows (arrays of cell values or cell objects)',
                        },
                        chartType: {
                          type: 'string',
                          description: 'For type=chart: Chart type (bar, line, pie, area, scatter, bubble, doughnut, radar)',
                        },
                        chartData: {
                          type: 'array',
                          description: 'For type=chart: Array of data series with name, labels, values',
                        },
                        path: {
                          type: 'string',
                          description: 'For type=image: Path to image file or base64 data URI',
                        },
                        shapeType: {
                          type: 'string',
                          description: 'For type=shape: Shape type (rectangle, ellipse, roundRectangle, triangle, etc.)',
                        },
                        options: {
                          type: 'object',
                          description: 'Element-specific options (positioning: x, y, w, h; formatting: fontSize, color, bold, etc.)',
                        },
                      },
                      required: ['type'],
                    },
                  },
                },
              },
            },
            filename: {
              type: 'string',
              description: 'Filename for the exported file (without extension)',
              default: 'output',
            },
            description: {
              type: 'string',
              description: 'Optional description of the file contents',
            },
            options: {
              type: 'object',
              description: 'Presentation options',
              properties: {
                layout: {
                  type: 'string',
                  enum: ['16x9', '16x10', '4x3'],
                  description: 'Slide layout/aspect ratio',
                  default: '16x9',
                },
                author: {
                  type: 'string',
                  description: 'Presentation author name',
                },
                title: {
                  type: 'string',
                  description: 'Presentation title',
                },
                subject: {
                  type: 'string',
                  description: 'Presentation subject',
                },
              },
            },
          },
          required: ['slides'],
        },
      },
    ],
  };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  if (name === 'pptx_export') {
    try {
      const {
        slides,
        filename = 'output',
        description,
        options = {},
      } = args;

      // Validate input
      if (!slides || !Array.isArray(slides)) {
        throw new Error('Slides must be provided as an array of slide objects');
      }

      if (slides.length === 0) {
        throw new Error('At least one slide must be provided');
      }

      // Create new PowerPoint presentation
      console.error('Creating PowerPoint presentation...');
      const pptx = new PptxGenJS();

      // Set presentation properties
      if (options.author) pptx.author = options.author;
      if (options.title) pptx.title = options.title;
      if (options.subject) pptx.subject = options.subject;

      // Set layout
      const layout = options.layout || '16x9';
      switch (layout) {
        case '16x9':
          pptx.layout = 'LAYOUT_16x9';
          break;
        case '16x10':
          pptx.layout = 'LAYOUT_16x10';
          break;
        case '4x3':
          pptx.layout = 'LAYOUT_4x3';
          break;
      }

      // Process each slide
      for (let i = 0; i < slides.length; i++) {
        console.error(`Processing slide ${i + 1} of ${slides.length}...`);

        const slideData = slides[i];
        const slide = pptx.addSlide();

        // Set slide background if specified
        if (slideData.backgroundColor) {
          slide.background = { color: slideData.backgroundColor };
        }

        // Process elements on the slide
        if (slideData.elements && Array.isArray(slideData.elements)) {
          for (const element of slideData.elements) {
            const elementType = element.type;
            const elementOptions = element.options || {};

            try {
              switch (elementType) {
                case 'text':
                  // Add text element
                  if (element.text) {
                    slide.addText(element.text, elementOptions);
                  }
                  break;

                case 'table':
                  // Add table element
                  if (element.rows && Array.isArray(element.rows)) {
                    slide.addTable(element.rows, elementOptions);
                  }
                  break;

                case 'chart':
                  // Add chart element
                  if (element.chartType && element.chartData) {
                    // Map chart type string to PptxGenJS chart type
                    const chartTypeMap = {
                      bar: pptx.ChartType.bar,
                      line: pptx.ChartType.line,
                      pie: pptx.ChartType.pie,
                      area: pptx.ChartType.area,
                      scatter: pptx.ChartType.scatter,
                      bubble: pptx.ChartType.bubble,
                      doughnut: pptx.ChartType.doughnut,
                      radar: pptx.ChartType.radar,
                      bar3d: pptx.ChartType.bar3D,
                    };

                    const chartType = chartTypeMap[element.chartType.toLowerCase()] || pptx.ChartType.bar;
                    slide.addChart(chartType, element.chartData, elementOptions);
                  }
                  break;

                case 'image':
                  // Add image element
                  if (element.path) {
                    slide.addImage({ path: element.path, ...elementOptions });
                  }
                  break;

                case 'shape':
                  // Add shape element
                  if (element.shapeType) {
                    // Map shape type string to PptxGenJS shape type
                    const shapeTypeMap = {
                      rectangle: pptx.ShapeType.rect,
                      ellipse: pptx.ShapeType.ellipse,
                      roundRectangle: pptx.ShapeType.roundRect,
                      triangle: pptx.ShapeType.triangle,
                      diamond: pptx.ShapeType.diamond,
                      pentagon: pptx.ShapeType.pentagon,
                      hexagon: pptx.ShapeType.hexagon,
                      octagon: pptx.ShapeType.octagon,
                      star: pptx.ShapeType.star,
                      arrow: pptx.ShapeType.rightArrow,
                    };

                    const shapeType = shapeTypeMap[element.shapeType.toLowerCase()] || pptx.ShapeType.rect;
                    slide.addShape(shapeType, elementOptions);
                  }
                  break;

                default:
                  console.error(`Unknown element type: ${elementType}`);
              }
            } catch (elementError) {
              console.error(`Error adding ${elementType} element:`, elementError);
              // Continue processing other elements
            }
          }
        }
      }

      // Generate UUID and filename
      const uuid = uuidv4();
      const sanitizedFilename = filename.replace(/[^a-z0-9_-]/gi, '_');
      const fullFilename = `${sanitizedFilename}_${uuid}.pptx`;

      // Generate PPTX to buffer (in-memory)
      console.error('Generating PowerPoint file in memory...');
      const fileBuffer = await pptx.write({ outputType: 'nodebuffer' });
      const fileSize = getFileSizeString(fileBuffer);

      // Write PPTX to file system
      const filepath = await writePPTXToFile(fileBuffer, fullFilename);

      console.error(`✅ PPTX generated: ${fullFilename} (${fileSize})`);
      console.error(`   Slides: ${slides.length}, Layout: ${layout}`);
      console.error(`   Saved to: ${filepath}`);

      // Return simplified response with essential information
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                path: fullFilename,
                filetype: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                filename: fullFilename,
                filesize: fileSize,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      console.error('Error processing PPTX export:', error);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: error.message || 'Unknown error',
              },
              null,
              2
            ),
          },
        ],
        isError: true,
      };
    }
  }

  throw new Error(`Unknown tool: ${name}`);
});

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('PPTX Export MCP Server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
