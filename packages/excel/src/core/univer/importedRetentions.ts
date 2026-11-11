// xlsx 导入条件格式/数据验证的落位：解析侧（retentionImport）产出纯数据，
// 这里在 Univer 挂载后经 facade / 注册命令写入，保证快照的 resources 里能正常序列化。
import { createCfId, iconMap, type IConditionFormattingRule, type IIconSetType } from '@univerjs/sheets-conditional-formatting';
import { getUniverApi } from './controller';
import {
  peekImportedRetention,
  shiftImportedRetention,
  unshiftImportedRetention,
  type ImportedCfRule,
  type ImportedDvRule,
  type ImportedIconSetDescriptor,
  type RetentionTask,
} from '../xlsx/retentionImport';

type UniverApi = NonNullable<ReturnType<typeof getUniverApi>>;
type WorkbookLike = NonNullable<ReturnType<UniverApi['getActiveWorkbook']>>;
type WorksheetLike = NonNullable<ReturnType<WorkbookLike['getSheetByName']>>;

let dvCounter = 0;
function dvUid(): string {
  dvCounter += 1;
  return `imported-dv-${Date.now().toString(36)}-${dvCounter}`;
}

function isIconSet(rule: ImportedCfRule['rule']): rule is ImportedIconSetDescriptor {
  return rule.type === 'iconSet' && Array.isArray((rule as ImportedIconSetDescriptor).values);
}

/** Excel 图标名称 → Univer 图标集的 iconId 列表（阈值按升序对应图标） */
function iconSetRule(descriptor: ImportedIconSetDescriptor): IConditionFormattingRule['rule'] {
  const iconType = (descriptor.iconType in iconMap ? descriptor.iconType : '3TrafficLights1') as IIconSetType;
  const icons = iconMap[iconType];
  const config = descriptor.values.map((value, index) => ({
    operator: 'greaterThanOrEqual' as never,
    value: { type: value.type, ...(value.value != null ? { value: value.value } : {}) } as never,
    iconType,
    iconId: icons?.[index] ?? '',
  }));
  return { type: 'iconSet', isShowValue: descriptor.isShowValue, config } as unknown as IConditionFormattingRule['rule'];
}

function applyCf(ws: WorksheetLike, task: ImportedCfRule): void {
  const rule = isIconSet(task.rule) ? iconSetRule(task.rule) : task.rule;
  ws.addConditionalFormattingRule({
    ranges: task.ranges,
    cfId: createCfId(),
    stopIfTrue: task.stopIfTrue,
    rule,
  } as IConditionFormattingRule);
}

function applyDv(api: UniverApi, wb: WorkbookLike, ws: WorksheetLike, task: ImportedDvRule): void {
  const { sheetName: _sheetName, ranges, ...rule } = task;
  void api.executeCommand('sheet.command.addDataValidation', {
    unitId: wb.getId(),
    subUnitId: ws.getSheetId(),
    rule: { ...rule, uid: dvUid(), ranges },
  } as never);
}

/** 消费导入队列：条件格式/数据验证逐条落位。任何一条失败只告警，不影响其余。返回成功落位条数 */
export function insertQueuedImportRetentions(): number {
  const api = getUniverApi();
  const wb = api?.getActiveWorkbook();
  if (!api || !wb) return 0;
  let applied = 0;
  for (;;) {
    if (getUniverApi() !== api) return applied;
    const task = peekImportedRetention();
    if (!task) return applied;
    shiftImportedRetention();
    try {
      applyOne(api, wb, task);
      applied += 1;
    } catch (err) {
      if (getUniverApi() !== api) {
        unshiftImportedRetention(task);
        return applied;
      }
      console.warn('[retentionImport] 条件格式/数据验证落位失败，已跳过', task.sheetName, err);
    }
  }
}

function applyOne(api: UniverApi, wb: WorkbookLike, task: RetentionTask): void {
  const ws = wb.getSheetByName(task.sheetName);
  if (!ws) return;
  if (task.kind === 'cf') applyCf(ws, task);
  else applyDv(api, wb, ws, task);
}
