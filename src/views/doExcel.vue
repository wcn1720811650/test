<script setup lang="ts">
/*! sheetjs (C) SheetJS -- https://sheetjs.com */
import {ref, unref, onMounted, defineComponent, computed, watch, watchEffect, reactive, warn} from "vue";
import VueTableLite from "vue3-table-lite/ts";
import {read, utils, WorkSheet, writeFile} from "xlsx";
import {useDraggable} from '@vueuse/core';
import {Alert} from "ant-design-vue";


type DataSet = { [index: string]: WorkSheet; };
type Row = any[];
type RowCB = (row: Row) => string;
type Column = { field: string; label: string; display: RowCB; };
type RowCol = { rows: Row[]; cols: Column[]; };

const open = ref(false)
const modalTitleRef = ref(null);
const showModal = () => {
  open.value = true;
};
const handleOk = e => {
  open.value = false;
  console.log(rows.value[0])
  console.log(rows.value[0][1])

  for (let i = 1; i < rows.value.length; i++) {
    if (rows.value[i][34]! < MinTime || rows.value! > MaxTime) {
      rows.value.splice(i, 1)

    }
  }
  console.log(rows.value.length)
};


const subOpen1 = ref<boolean>(false);
const showSubModal1 = () => {
  subOpen1.value = true;
};
const handleSubOk1 = (e: MouseEvent) => {
  subOpen1.value = false;

};


const subOpen2 = ref<boolean>(false);
const showSubModal2 = () => {
  subOpen1.value = true;
};
const handleSubOk2 = (e: MouseEvent) => {
  subOpen1.value = false;
};


const subOpen3 = ref<boolean>(false);
const showSubModal3 = () => {
  subOpen1.value = true;
};
const handleSubOk3 = (e: MouseEvent) => {
  subOpen1.value = false;
};


const currFileName = ref<string>("");
const currSheet = ref<string>("");
const sheets = ref<string[]>([]);
const workBook = ref<DataSet>({} as DataSet);
const rows = ref([[]]);
const columns = ref();
const MaxTime = ref<number>(0);
const MinTime = ref<number>(0);
const ForcedTitle = ref<number>(0)
const ForcedAnswer = ref<number>(0)
const ForcedItemArr = ref([])
const repeatValue1 = ref<number>(0)
const repeatValue2 = ref<number>(0)
const RepeatItem = ref([])
const hasCheckbox = ref(true)
const selectedRowKeys = ref()
const selectedRows = ref()
const readColumns = [{title: 'title', width: 1000, dataIndex: 'title', key: '1'}]


const dataResource = []
const importExcel = () => {
  for (let i = 0; i < rows.value[0].length; i++) {
    dataResource.push({
      key: i,
      title: rows.value[0][i],
      width: 1000
    })
  }
}

const rowSelection = {
  onChange: (selectedRowKeys, selectedRows) => {
    console.log(`selectedRowKeys: ${selectedRowKeys}`, 'selectedRows: ', selectedRows);
  },
  getCheckboxProps: record => ({
    disabled: record.name === 'Disabled User',
    // Column configuration not to be checked
    name: record.name,
  }),
};

const exportTypes: string[] = ["xlsx", "xlsb", "csv", "html"];
const addForcedTopic = () => {
  if (ForcedTitle.value == 0 || ForcedAnswer.value == 0) {
    window.alert("强制题编号和答案不能为0！")
  }

  for (let i = 0; i < ForcedItemArr.value.length; i++) {
    if (ForcedItemArr.value[i].topic.value == ForcedTitle.value) {
      window.alert("您已经添加过此题目了")
      return;
    }
  }
  ForcedItemArr.value.push({topic: ForcedTitle.value, answer: ForcedAnswer.value})
  ForcedTitle.value = 0;
  ForcedAnswer.value = 0;
}
const addRepeatTopic = () => {
  if (repeatValue1.value !== repeatValue2.value || repeatValue1.value === 0 || repeatValue2.value === 0) {
    window.alert("重复题编号必须相同且值不为0！")
  }
  RepeatItem.value.push({repeatTopic1: repeatValue1.value, repeatTopic2: repeatValue2.value})
}
let cell = 0;

function resetCell() {
  cell = 0;
}

const getRowsCols = (data: DataSet, sheetName: string): RowCol => ({
  rows: utils.sheet_to_json<Row>(data[sheetName], {header: 1}),
  cols: Array.from({
    length: utils.decode_range(data[sheetName]["!ref"] || "A1").e.c + 1
  }, (_, i) => (<Column>{field: String(i), label: utils.encode_col(i), display: makeDisplay(i)}))
});

const makeDisplay = (col: number): RowCB => (row: Row) => `<span
  style="user-select: none; display: block"
  onblur="endEdit(event)" ondblclick="startEdit(event)"
  position="${Math.floor(cell++ / columns.value.length)}.${col}"
  onkeydown="endEdit(event)">${row?.[col] ?? "&nbsp;"}</span>`;

(window as any).startEdit = function (ev: MouseEvent) {
  (ev?.target as HTMLSpanElement).contentEditable = "true";
  (ev?.target as HTMLSpanElement).focus();
};

(window as any).endEdit = function (ev: FocusEvent | KeyboardEvent) {
  if (typeof (ev as KeyboardEvent).key == "undefined" || (ev as KeyboardEvent).key === "Enter") {
    const pos = (ev.target as HTMLSpanElement)?.getAttribute("position")?.split(".");
    if (!pos) return;

    (ev?.target as HTMLSpanElement).contentEditable = "true";

    rows.value[+pos[0]][+pos[1]] = (ev.target as HTMLSpanElement).innerText;

    workBook.value[currSheet.value] = utils.json_to_sheet(rows.value, {
      header: columns.value.map((col: Column) => col.field),
      skipHeader: true,
    });
  }
};

async function importAB(ab: ArrayBuffer, name: string): Promise<void> {
  const data = read(ab);
  currFileName.value = name;
  currSheet.value = data.SheetNames?.[0];
  sheets.value = data.SheetNames;
  workBook.value = data.Sheets;
  selectSheet(currSheet.value);
}

async function importFile(ev: Event): Promise<void> {
  const file = (ev.target as HTMLInputElement)?.files?.[0];
  if (!file) return;
  await importAB(await file.arrayBuffer(), file.name);
  importExcel()
}

function exportFile(type: string): void {
  const wb = utils.book_new();

  sheets.value.forEach((sheet) => {
    utils.book_append_sheet(wb, workBook.value[sheet], sheet);
  });

  writeFile(wb, `sheet.${type}`);
}

function selectSheet(sheet: string): void {
  const {rows: newRows, cols: newCols} = getRowsCols(workBook.value, sheet);
  resetCell();
  rows.value = newRows;
  columns.value = newCols;
  currSheet.value = sheet;

}

/* Download from https://sheetjs.com/pres.numbers */
onMounted(async () => {
  const response = await fetch("");
  await importAB(await response.arrayBuffer(), "pres.numbers");
});
</script>

<template>
  <header class="imp-exp">
    <div class="import">
      <input type="file" id="import" @change="importFile"/>
      <label for="import">import</label>
    </div>
    <span v-if="currFileName">{{ currFileName }}</span>
    <div class="export" v-if="currFileName">
      <span>export</span>
      <ul>
        <li v-for="(type, idx) in exportTypes" :key="idx" @click="exportFile(type)">
          {{ `.${type}` }}
        </li>
      </ul>
    </div>
  </header>
  <div class="sheets">
    <span
        v-for="(sheet, idx) in sheets"
        :key="idx"
        @click="selectSheet(sheet)"
        :class="[currSheet === sheet ? 'selected' : '']"
    >
      {{ sheet }}
    </span>
  </div>
  <div>
    <a-button type="normal" @click="showModal">筛选</a-button>
  </div>

  <!--  <vue-table-lite :has-checkbox="hasCheckbox" :is-static-mode="true" :page-size="50" :columns="columns" :rows="rows"></vue-table-lite>-->

  <a-table :columns="readColumns" :data-source="dataResource" :row-selection="rowSelection"/>

  <a-modal ref="modalRef" v-model:open="open" :wrap-style="{ overflow: 'hidden' }" @ok="handleOk">
    <div style="display: flex;flex-direction: column">
      <div class="contentStyle">
        作答时间上限
        当前值：{{ MaxTime + "s" }}
        <a-input-number id="inputNumber" v-model:value="MaxTime" size="small" :min="MinTime+1"/>
      </div>
      <div class="contentStyle">
        作答时间下限
        当前值：{{ MinTime + "s" }}
        <a-input-number id="inputNumber" v-model:value="MinTime" size="small" :min="0" :max="MaxTime-1"/>
      </div>
      <div class="contentStyle" style="display: flex;flex-direction: row;justify-content: space-between">
        <div>
          <span>强制题编号</span>
          <a-input-number id="inputNumber" v-model:value="ForcedTitle" size="small" :min="0"/>
        </div>
        <div>
          <span>答案</span>
          <a-input-number id="inputNumber" v-model:value="ForcedAnswer" size="small" :min="0"/>
        </div>
        <a-button type="primary" size="small" class="selectBtn" @click="addForcedTopic">添加</a-button>
      </div>
      <div class="contentStyle" style="display: flex;flex-direction: row;justify-content: space-between">
        <div>
          <span>重复题编号</span>
          <a-input-number id="inputNumber" v-model:value="repeatValue1" size="small" :min="0"/>
          <a-input-number id="inputNumber" v-model:value="repeatValue2" size="small" :min="0"/>
          <a-button type="primary" size="small" class="selectBtn" @click="addRepeatTopic">添加</a-button>
        </div>
      </div>
      <!--      <div class="contentStyle" style="display: flex;flex-direction: row;justify-content: space-between">-->
      <!--        <div>1.请选择问卷中表示”序号“的一列</div>-->
      <!--        <a-button type="primary" size="small" class="selectBtn" @click="showSubModal1">选择</a-button>-->
      <!--      </div>-->
      <!--      <div class="contentStyle" style="display: flex;flex-direction: row;justify-content: space-between">-->
      <!--        <div>2.请选择问卷中表示”作答时间“的一列</div>-->
      <!--        <a-button type="primary" size="small" class="selectBtn" @click="showSubModal2">选择</a-button>-->
      <!--      </div>-->
      <!--      <div class="contentStyle" style="display: flex;flex-direction: row;justify-content: space-between">-->
      <!--        <div>3.请选择所有待筛选问题列</div>-->
      <!--        <a-button type="primary" size="small" class="selectBtn" @click="showSubModal3">选择</a-button>-->
      <!--      </div>-->
    </div>
    <template #title>
      <div ref="modalTitleRef" style="width: 100%; cursor: move">筛选</div>
    </template>
    <template #modalRender="{ originVNode }">
      <div :style="transformStyle">
        <component :is="originVNode"/>
      </div>
    </template>
  </a-modal>
  <div>
    <a-modal v-model:open="subOpen1" title="Basic Modal" @ok="handleSubOk1">
      <div v-for="(item,index) in rows[0] " :key="index" style="display: flex;flex-direction: row">
        <a-checkbox style="margin:0 10px 0 10px"></a-checkbox>
        <p>{{ index + 1 }}.{{ item }}</p>
      </div>
    </a-modal>
  </div>
  <div>
    <a-modal v-model:open="subOpen2" title="Basic Modal" @ok="handleSubOk2">
      <p>Some contents...</p>
      <p>Some contents...</p>
      <p>Some contents...</p>
    </a-modal>
  </div>
  <div>
    <a-modal v-model:open="subOpen3" title="Basic Modal" @ok="handleSubOk3">
      <p>Some contents...</p>
      <p>Some contents...</p>
      <p>Some contents...</p>
    </a-modal>
  </div>
</template>

<style>
.imp-exp {
  display: flex;
  justify-content: space-between;
  padding: 0.5rem;
  font-family: mono;
  color: #212529;
}

.import {
  font-size: medium;
}

.import input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
}

.import label {
  background-color: white;
  border: 1px solid;
  padding: 0.3rem;
}

.export:hover {
  border-bottom: none;
}

.export:hover ul {
  display: block;
}

.export span {
  padding: 0.3rem;
  border: 1px solid;
  cursor: pointer;
}

.export ul {
  display: none;
  position: absolute;
  z-index: 5;
  background-color: white;
  list-style: none;
  padding: 0.3rem;
  border: 1px solid;
  margin-top: 0.3rem;
  border-top: none;
}

.export ul li {
  padding: 0.3rem;
  text-align: center;
}

.export ul li:hover {
  background-color: lightgray;
  cursor: pointer;
}

.sheets {
  display: flex;
  justify-content: center;
  margin: 0.3rem;
  color: #212529;
}

.sheets span {
  border: 1px solid;
  padding: 0.5rem;
  margin: 0.3rem;
}

.sheets span:hover:not(.selected) {
  background-color: lightgray;
  cursor: pointer;
}

.selected {
  background-color: #343a40;
  color: white;
}

.selectBtn {
}

.contentStyle {
  margin: 20px 0 20px 20px;
}
</style>