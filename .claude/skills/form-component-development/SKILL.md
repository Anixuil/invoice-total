---
name: form-component-development
description: 为 itom-ui 的可视化表单设计器新增或修改表单组件，覆盖组件配置、拖拽设计器、属性面板、运行时预览、校验、插槽、动态数据和 Vue 代码生成。用户提到增加表单组件、扩充组件库、修改组件属性或修复表单组件渲染时使用此技能。
---

# 表单组件开发

## 目标

在当前项目中维护一套完整的表单组件协议，确保组件能被拖入设计器、保存为 `formContent`、在表单预览中正常工作，并在需要时正确导出 Vue 文件。

## 开始前

1. 读取项目根目录和目标目录下的 `AGENTS.md`，遵循其中关于验证、Git 和文件修改的约束。
2. 先查看与新组件最接近的现有组件，不要凭空设计另一套配置格式。重点参考：
   - `src/utils/generator/config.js`
   - `src/views/tool/build/RightPanel.vue`
   - `src/utils/generator/render.js`
   - `src/components/render/render.js`
   - `src/utils/generator/html.js`
   - `src/utils/generator/js.js`
3. 判断需求属于哪一类：
   - 仅新增 Element Plus 原生组件
   - 新增带选项或插槽的组件
   - 新增项目自定义业务组件
   - 新增布局容器组件
   - 修改已有组件属性、校验或导出行为

## 配置协议

在 `src/utils/generator/config.js` 中添加组件描述。至少明确：

```js
{
  label: '组件名称',
  tag: '组件注册名',
  tagIcon: '图标名',
  defaultValue: '',
  span: 24,
  required: false,
  regList: [],
  changeTag: true
}
```

约定：

- `tag` 必须是运行时可解析的组件名。
- 需要双向绑定的字段必须有 `defaultValue`，拖入后由设计器生成 `vModel`。
- 可编辑属性必须在默认配置中预先声明，右侧面板依赖“属性是否存在”决定是否显示编辑项。
- 选择类组件使用 `options`；级联类组件还要维护 `props`、`labelKey`、`valueKey`、`childrenKey` 和 `dataType`。
- 容器使用 `layout: 'rowFormItem'`、`children: []`；普通字段使用 `layout: 'colFormItem'`。
- 不复用已有组件的 `vModel`，字段名必须能通过保存前的重复检查。

## 必须覆盖的链路

### 1. 组件库与设计器

修改 `config.js` 后确认组件出现在正确的左侧分组。设计器位于 `src/views/tool/build/index.vue`，拖入时会复制配置并生成：

- `formId`：组件标识
- `renderKey`：列表渲染 key
- `vModel`：表单模型字段名

如果是自定义拖拽或嵌套容器，检查 `DraggableItem.vue` 和 `index.vue` 中的 `cloneComponent`、`createIdAndKey`、递归 `children` 处理。

### 2. 右侧属性面板

在 `src/views/tool/build/RightPanel.vue` 增加新属性的编辑控件。使用与字段属性一致的条件判断，例如：

```vue
<el-form-item v-if="activeData.max !== undefined" label="最大值">
  <el-input-number v-model="activeData.max" />
</el-form-item>
```

涉及属性联动时同步处理现有的 change 方法，例如日期类型、选项类型、范围模式、多选模式和颜色格式。

### 3. 设计器画布渲染

设计器使用 `src/utils/generator/render.js`。确认它能：

- 将配置属性传给组件
- 使用 `modelValue` 与 `update:modelValue` 完成双向绑定
- 处理选项子组件或上传按钮等特殊插槽
- 忽略 `layout`、`vModel`、`defaultValue` 等内部配置，不把它们错误透传给组件

### 4. 保存后预览渲染

表单预览使用 `src/components/render/render.js`，不是设计器的同一个渲染器。新增组件必须同时检查这条链路。

自定义组件优先在 `src/main.js` 全局注册；如果组件有复杂插槽，在 `src/components/render/slots/` 增加以 `tag` 命名的插槽适配文件。组件应支持：

```vue
<MyComponent
  :model-value="value"
  @update:model-value="value = $event"
/>
```

检查 `Parser.vue` 的字段读取和回写同时兼容新旧格式：

```js
const vModel = scheme.__vModel__ || scheme.vModel
```

不要只使用 `__vModel__`，因为当前设计器保存的是扁平 `vModel` 结构。

### 5. 校验与动态数据

需要校验时：

- 在组件配置中定义 `required` 和 `regList`。
- 在 `src/utils/generator/config.js` 的 `trigger` 中增加组件触发方式。
- 确认 `src/utils/generator/js.js` 能生成规则。
- 确认 `src/components/parser/Parser.vue` 能将规则应用到预览表单。

需要远程选项时，沿用 `dataType: 'dynamic'`、`method`、`url`、`dataPath`、`dataConsumer` 的现有模式，并检查 `tool/build/index.vue` 的 `fetchData`、`setRespData`。

### 6. Vue 代码导出

只有在用户要求支持“导出 Vue 文件”时才补生成器：

- `src/utils/generator/html.js`：生成组件标签、属性、子项和插槽。
- `src/utils/generator/js.js`：生成默认模型、校验规则、选项、请求方法、上传逻辑。
- `src/utils/generator/css.js`：需要专属样式时增加样式生成。

设计器能渲染不代表导出的文件能运行；必须分别检查两条链路。

## 自定义业务组件

新增项目组件时按以下顺序处理：

1. 编写组件并明确 `modelValue` / `update:modelValue` 协议。
2. 在 `src/main.js` 注册组件，或在动态渲染器中增加显式映射。
3. 在 `config.js` 增加设计器配置。
4. 在 `RightPanel.vue` 增加可编辑属性。
5. 在两个渲染器中验证属性和事件传递。
6. 根据需求补插槽、校验、动态数据和代码生成。
7. 如果使用 `tagIcon`，补齐对应 SVG 图标资源和 `src/utils/generator/icon.json` 配置。

## 交付前静态检查

只检查本次修改涉及的文件，至少确认：

- `tag`、全局注册名和实际组件名一致。
- 设计器预览与保存后 `Parser` 预览都能渲染。
- `modelValue` 的初始值、输入回写和重置行为一致。
- `vModel` 唯一且不会覆盖其他字段。
- 选项、插槽、上传、动态请求和校验没有遗漏。
- 旧版 `__config__` / `__vModel__` 结构仍可兼容时，不破坏兼容逻辑。
- 遵循项目指令：除非用户明确要求，不运行构建、测试、lint 或类型检查；不执行 commit、worktree 或 push。
