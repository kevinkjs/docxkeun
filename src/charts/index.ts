/**
 * Charts (word/charts/chartN.xml, root c:chartSpace): read the cached data into a model, describe it
 * in one line of text, and draw it into display items for the SVG viewer and PDF export; write new
 * charts from a table of data, read any chart's data as that table, and patch it with new data.
 */
export { readChart, describeChart, autoColor } from './model.js';
export type {
  ChartModel,
  ChartPlot,
  ChartSeries,
  ChartAxis,
  ChartLegend,
  ChartText,
  ChartPoint,
  DataLabels,
  DataLabelFlags,
  LineStyle,
  ManualLayout,
  MarkerStyle,
  MarkerSymbol,
  NumFmt,
  PlotType,
  Grouping,
  TextStyle,
  ReadChartOptions,
} from './model.js';
export { newChartSpace, readChartTable, patchChartSpace, validateChartData, ChartDataError } from './write.js';
export type { NewChart, NewChartType, LegendPosition, ChartChanges, ChartTable, ChartSeriesData } from './write.js';
export { renderChart, drawChart } from './render.js';
export type { ChartFrame, RenderChartOptions } from './render.js';
export { formatNumber, formatGeneral, isDateCode } from './format.js';
export { resolveColorElement, colorIn, fillOf, lineOf } from './color.js';
export type { Paint, SchemeColorFn, LineProps } from './color.js';
