/** The chart writer's types and ChartDataError are exported from the editor entry point. */
import { expect, it } from 'vitest';
import { ChartDataError, type ChartChanges, type ChartSeriesData, type ChartTable, type LegendPosition, type NewChart, type NewChartType } from '../src/api/index.js';

it('exports the chart types and ChartDataError', () => {
  const type: NewChartType = 'column';
  const legend: LegendPosition = 'top';
  const series: ChartSeriesData[] = [{ name: 'A', values: [1, 2] }];
  const chart: NewChart = { type, legend, categories: ['x', 'y'], series };
  const changes: ChartChanges = { title: 'T' };
  const table: Pick<ChartTable, 'kind'> = { kind: 'clustered column' };
  expect([chart.type, changes.title, table.kind]).toEqual(['column', 'T', 'clustered column']);
  expect(new ChartDataError('x')).toBeInstanceOf(Error);
});
