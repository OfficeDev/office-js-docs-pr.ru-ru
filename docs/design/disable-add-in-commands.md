---
title: Включение и отключение команд надстроек
description: Узнайте, как изменить состояние ("Включено" или "Отключено") настраиваемых кнопок ленты и элементов меню в веб-надстройке Office.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: fa4830c0112486bbad7a13edf78e0c8c4277e143
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217895"
---
# <a name="enable-and-disable-add-in-commands"></a>Включение и отключение команд надстроек

Если некоторые функции надстройки должны быть доступны только в определенном контексте, вы можете включить или отключить настраиваемые команды надстройки программными средствами. Например, функция, изменяющая заголовок таблицы, должна быть включена, только когда курсор находится в таблице.

Также можно указать, будет ли команда включена или отключена при открытии ведущего приложения Office.

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> [Основные концепции команд надстроек](add-in-commands.md)

## <a name="rules-and-gotchas"></a>Правила и подсказки

### <a name="single-line-ribbon-in-office-on-the-web"></a>Однострочная лента в Office в Интернете

В Office в Интернете API и разметка манифеста, описанные в этой статье, применимы только к однострочной ленте. Они не оказывают влияния на многострочную ленту. Они затрагивают обе ленты в классических приложениях Office. Дополнительные сведения об обеих лентах см. в статье [Использование упрощенной ленты](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).

### <a name="shared-runtime-required"></a>Необходима общая среда выполнения

API и разметка манифеста надстройки, описанные в этой статье, требуют использования общей среды выполнения. Для этого выполните следующие действия.

1. В элементе манифеста [Runtimes](../reference/manifest/runtimes.md) добавьте следующий дочерний элемент: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (Если в манифесте еще нет элемента `<Runtimes>`, создайте его в качестве первого дочернего элемента `<Host>` в разделе `VersionOverrides`.)
2. В разделе [Resources.Urls](../reference/manifest/resources.md) манифеста добавьте следующий дочерний элемент:`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, где `{MyDomain}` домен надстройки и `{path-to-start-page}`путь к начальной странице надстройки; например: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. В зависимости от того, есть ли в вашей надстройке область задач, файл функций или настраиваемая функция Excel, необходимо выполнить одно или несколько из описанных ниже трех действий.

    - Если в надстройке есть область задач, задайте атрибут `resid` элемента [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md)точно так же, как в `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если в настройке есть настраиваемая функция Excel, установите атрибут `resid` элемента [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если в настройке есть файл функций, установите атрибут `resid` элемента [FunctionFile](../reference/manifest/functionfile.md) точно так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Установка состояния "Отключено" по умолчанию

По умолчанию при запуске приложения Office любая команда надстройки включается. Если вы хотите, чтобы при запуске приложения Office настраиваемая кнопка или элемент меню были отключены, укажите это в манифесте. Просто добавьте элемент [Enabled](../reference/manifest/enabled.md) (со значением`false`) сразу *под* (не внутри) элемента [Action](../reference/manifest/action.md) в объявлении элемента управления. Ниже показана базовая структура.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Изменение состояния программными средствами

Ниже приведены основные действия по изменению состояния "Включено" команды надстройки.

1. Создайте объект [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata), в котором (1) указаны идентификаторы команды и ее родительской вкладки в соответствии с манифестом и (2) указано состояние команды ("Включено" или "Отключено").
2. Перенесите объект **RibbonUpdaterData** в метод [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-).

Ниже приведен простой пример. Обратите внимание, что "MyButton" и "OfficeAddinTab1" скопированы из манифеста.

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "OfficeAppTab1", 
                controls: [
                {
                    id: "MyButton", 
                    enabled: true
                }
            ]}
        ]});
}
```

Кроме того, мы предоставляем несколько интерфейсов (типов) для упрощения создания объекта **RibbonUpdateData**. Ниже приводится аналогичный пример в TypeScript, в котором используются эти типы.

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Office определяет время обновления состояния ленты. Метод **requestUpdate()** ставит запрос на обновление в очередь. Этот метод устранит объект Promise, как только он поставит запрос в очередь, а не при обновлении ленты.

## <a name="change-the-state-in-response-to-an-event"></a>Изменение состояния в ответ на событие

Обычно состояние ленты необходимо изменить, когда инициированное пользователем событие изменяет контекст надстройки.

Рассмотрим сценарий, в котором кнопка должна быть включена, только когда активирована диаграмма. Во-первых, задайте значение `false` для элемента [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте. Пример см. выше.

Во-вторых, назначьте обработчиков. Это обычно выполняется с помощью метода **Office.onReady**, как в приведенном ниже примере, где обработчики (созданные позднее) назначаются событиям **onActivated** и **onDeactivated** всех диаграмм на листе.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

В-третьих, определите обработчик `enableChartFormat`. Ниже приведен простой пример. Более надежный способ изменения состояния элемента управления см. в разделе **Рекомендация: проверка на наличие ошибок в состоянии элементов управления** ниже.

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

В-четвертых, определите обработчик `disableChartFormat`. Он будет идентичен `enableChartFormat`, только для свойства объекта кнопки **enabled** будет задано значение `false`.

## <a name="best-practice-test-for-control-status-errors"></a>Рекомендация: проверка на наличие ошибок в состоянии элементов управления

В некоторых случаях после вызова `requestUpdate` лента не обновляется, поэтому гиперсостояние элемента управления не изменяется. По этой причине рекомендуется отслеживать состояние элементов управления надстройки. Надстройка должна соответствовать приведенным ниже требованиям.

1. При вызове `requestUpdate` в коде указывается предполагаемое состояние настраиваемых кнопок и элементов меню.
2. При щелчке пользовательского элемента управления первый код в обработчике проверяет, должна ли кнопка быть интерактивной. Если нет, код сообщит об ошибке или запишет ее в журнал и попробует еще раз установить для кнопок предполагаемое состояние.

В приведенном ниже примере показана функция, с помощью которой можно отключить кнопку и записать ее состояние. Обратите внимание, что `chartFormatButtonEnabled` — глобальная логическая переменная, которая инициализируется до того же значения, что и элемент [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте.

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

В приведенном ниже примере показано, как обработчик кнопки проверяет ее на наличие неправильного состояния. Обратите внимание, что `reportError` — это функция, которая отображает или записывает в журнал ошибку.

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a>Обработка ошибок

В некоторых случаях Office не может обновить ленту и возвращает ошибку. Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть. Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`. Ниже приведен пример обработки этой ошибки. В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.

```javascript
function disableChartFormat() {
    try {
        var button = {id: "ChartFormatButton", enabled: false};
        var parentTab = {id: "CustomChartTab", controls: [button]};
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```

## <a name="test-for-platform-support-with-requirement-sets"></a>Тестирование поддержки платформ с использованием наборов обязательных элементов

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../develop/office-versions-and-requirement-sets.md).

Для включения или отключения API требуется поддержка следующего набора обязательных элементов:

- [AddinCommands 1.1](../reference/requirement-sets/add-in-commands-requirement-sets.md)
