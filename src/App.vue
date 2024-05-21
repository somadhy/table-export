<template>
  <DxDataGrid
    :data-source="dataSource"
    :remote-operations="false"
    :allow-column-reordering="true"
    :row-alternation-enabled="true"
    :show-borders="true"
    :width="'100%'"
    @content-ready="onContentReady"
    @exporting="onExporting"
  >
    <DxExport :enabled="true" />
    <DxColumn :group-index="0" data-field="Product" />
    <DxColumn
      data-field="Amount"
      caption="Sale Amount"
      data-type="number"
      format="currency"
      alignment="right"
    />
    <DxColumn
      :allow-grouping="false"
      data-field="Discount"
      caption="Discount %"
      data-type="number"
      format="percent"
      alignment="right"
      cell-template="discountCellTemplate"
      css-class="bullet"
    />
    <DxColumn data-field="SaleDate" data-type="date" />
    <DxColumn data-field="Region" data-type="string" />
    <DxColumn data-field="Sector" data-type="string" />
    <DxColumn data-field="Channel" data-type="string" />
    <DxColumn :width="150" data-field="Customer" data-type="string" />

    <DxGroupPanel :visible="true" />
    <DxSearchPanel :visible="true" :highlight-case-sensitive="true" />
    <DxGrouping :auto-expand-all="false" />
    <DxPager :allowed-page-sizes="pageSizes" :show-page-size-selector="true" />
    <DxPaging :page-size="10" />
    <template #discountCellTemplate="{ data: cellData }">
      <DiscountCell :cell-data="cellData" />
    </template>
  </DxDataGrid>
</template>
<script setup lang="ts">
import 'devextreme/dist/css/dx.light.css';
import 'devextreme/data/odata/store';

import {
  DxColumn,
  DxDataGrid,
  DxDataGridTypes,
  DxExport,
  DxGrouping,
  DxGroupPanel,
  DxPager,
  DxPaging,
  DxSearchPanel,
} from 'devextreme-vue/data-grid';
import DataSource from 'devextreme/data/data_source';
import { exportDataGrid } from 'devextreme/excel_exporter';
import { Workbook } from 'exceljs';
import saveAs from 'file-saver';

import DiscountCell from './components/DiscountCell.vue';

const dataSource = new DataSource({
  store: {
    type: 'odata',
    version: 2,
    url: 'https://js.devexpress.com/Demos/SalesViewer/odata/DaySaleDtoes',
    key: 'Id',
    beforeSend(request) {
      const year = new Date().getFullYear() - 1
      request.params.startDate = `${year}-05-10`
      request.params.endDate = `${year}-5-15`
    }
  }
})

const pageSizes = [10, 25, 50, 100]

let collapsed = false

const onContentReady = (e: DxDataGridTypes.ContentReadyEvent) => {
  if (!collapsed) {
    e.component.expandRow(['EnviroCare'])
    collapsed = true
  }
}

const onExporting = (e) => {
  const workbook = new Workbook()
  const worksheet = workbook.addWorksheet('Companies')
  const groupRows = new Array()

  const topLeftRow = 2
  const topLeftCol = 2

  exportDataGrid({
    autoFilterEnabled: true,
    component: e.component,
    worksheet: worksheet,
    topLeftCell: { row: topLeftRow, column: topLeftCol },
    customizeCell: function (options) {
      const { gridCell, excelCell } = options
      if (gridCell.rowType === 'data' && gridCell.column.dataField === 'Customer') {
        excelCell.font = { color: { argb: 'FF0000FF' }, underline: true }
        excelCell.alignment = { horizontal: 'left' }
      } else if (gridCell.rowType == 'group' && gridCell.value) {
        // Пополнение массива строк для объединения.
        groupRows.push(excelCell.row)

        // Оформление ячейки - заголовка группы.
        excelCell.font = { color: { argb: 'FFFF0000' }, fontName: 'Tahoma', bold: true, size: 14 }
        excelCell.alignment = { horizontal: 'center' }
      } else if (gridCell.rowType == 'header') {
        // Ячейки заголовка.
        excelCell.font = { color: { argb: 'FFFFFFFF' }, fontName: 'Tahoma', bold: true, size: 14 }
        excelCell.fill = {
          type: 'gradient',
          gradient: 'path',
          center: { left: 0.5, top: 0.5 },
          stops: [
            { position: 0, color: { argb: 'FFFF0000' } },
            { position: 1, color: { argb: 'FF00FF00' } }
          ]
        }
        excelCell.alignment = { horizontal: 'center' }
      }

      // Бордюры
      excelCell.border = {
        top: { style: 'double', color: { argb: 'FF00FF00' } },
        left: { style: 'double', color: { argb: 'FF00FF00' } },
        bottom: { style: 'double', color: { argb: 'FF00FF00' } },
        right: { style: 'double', color: { argb: 'FF00FF00' } }
      }
    }
  }).then(function () {
    // Нужно перебрать все столбцы требующие изменения ширины.
    worksheet.columns.forEach((column) => {
      if (column.letter !== 'A') column.width = 40
    })

    // Требуется заранее собрать список строк для которых требуется объединение.
    groupRows.forEach((row) => {
      // Объединение строк
      worksheet.mergeCells(row, topLeftCol, row, worksheet.columns.length)
    })

    workbook.xlsx.writeBuffer().then(function (buffer) {
      saveAs(new Blob([buffer], { type: 'application/octet-stream' }), 'Companies.xlsx')
    })
  })
}
</script>
