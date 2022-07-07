---
title: Включение и отключение команд надстроек
description: Узнайте, как изменить состояние ("Включено" или "Отключено") настраиваемых кнопок ленты и элементов меню в веб-надстройке Office.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca9e35026acb91a54affa8215178f2eaa6cbd4c9
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659824"
---
# <a name="enable-and-disable-add-in-commands"></a>Включение и отключение команд надстроек

Если некоторые функции надстройки должны быть доступны только в определенном контексте, вы можете включить или отключить настраиваемые команды надстройки программными средствами. Например, функция, изменяющая заголовок таблицы, должна быть включена, только когда курсор находится в таблице.

Можно также указать, включена или отключена команда при запуске клиентского приложения Office.

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Только поддержка приложений и платформ Office

API, описанные в этой статье, доступны только в Excel, PowerPoint и Word.

### <a name="test-for-platform-support-with-requirement-sets"></a>Тестирование поддержки платформ с использованием наборов обязательных элементов

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы обязательных элементов, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли сочетание приложений и платформ Office API, необходимые надстройке. Дополнительные сведения см. в статьях [о версиях Office и наборах обязательных элементов](../develop/office-versions-and-requirement-sets.md).

API включения и отключения относятся к набору обязательных элементов [RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) .

> [!NOTE]
> Набор **обязательных элементов RibbonApi 1.1** пока не поддерживается в манифесте, поэтому его нельзя указать в разделе манифеста **\<Requirements\>** . Чтобы проверить поддержку, код должен вызвать .`Office.context.requirements.isSetSupported('RibbonApi', '1.1')` Если и *только в том* случае, если этот вызов возвращается `true`, код может вызывать API включения и отключения. Если вызов возвращается `isSetSupported` `false`, все пользовательские команды надстройки будут включены все время. Необходимо разработать рабочую надстройку и все инструкции в приложении, чтобы указать, как она будет работать, если набор обязательных элементов **RibbonApi 1.1** не поддерживается. Дополнительные сведения и примеры `isSetSupported`использования см. в статье "Указание приложений [Office и требований К API](../develop/specify-office-hosts-and-api-requirements.md)", в частности проверки среды выполнения на наличие поддержки методов [и наборов обязательных элементов](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (Раздел " [Указание версий и платформ Office](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) " для размещения надстройки этой статьи не относится к ленте 1.1.)

## <a name="shared-runtime-required"></a>Необходима общая среда выполнения

API и разметка манифеста надстройки, описанные в этой статье, требуют использования общей среды выполнения. Для этого выполните следующие действия.

1. В элементе манифеста [Runtimes](/javascript/api/manifest/runtimes) добавьте следующий дочерний элемент: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (Если в манифесте **\<Runtimes\>** еще нет элемента, **\<Host\>** создайте его в качестве первого дочернего элемента в разделе **\<VersionOverrides\>** .)
2. В разделе [Resources.Urls](/javascript/api/manifest/resources) манифеста добавьте следующий дочерний элемент:`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, где `{MyDomain}` домен надстройки и `{path-to-start-page}`путь к начальной странице надстройки; например: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. В зависимости от того, содержит ли надстройка область задач, файл функции или пользовательскую функцию Excel, необходимо выполнить одно или несколько из следующих трех действий.

    - Если надстройка содержит область задач, задайте `resid` атрибут [действия](/javascript/api/manifest/action).[ Для элемента SourceLocation](/javascript/api/manifest/sourcelocation) используется точно та же `resid` **\<Runtime\>** строка, что и для элемента на шаге 1, `Contoso.SharedRuntime.Url`например . Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если надстройка содержит пользовательскую функцию Excel, задайте `resid` атрибут [страницы](/javascript/api/manifest/page).[ Элемент SourceLocation](/javascript/api/manifest/sourcelocation) точно такой же строки, как вы использовали `resid` **\<Runtime\>** для элемента на шаге 1, например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если надстройка содержит файл функции, задайте для атрибута элемента [FunctionFile](/javascript/api/manifest/functionfile) `resid` **\<Runtime\>** ту же строку, `resid` что и для элемента на шаге 1, `Contoso.SharedRuntime.Url`например . Элемент должен выглядеть следующим образом:`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Установка состояния "Отключено" по умолчанию

По умолчанию при запуске приложения Office любая команда надстройки включается. Если вы хотите, чтобы при запуске приложения Office настраиваемая кнопка или элемент меню были отключены, укажите это в манифесте. Просто добавьте элемент [Enabled](/javascript/api/manifest/enabled) (со значением`false`) сразу *под* (не внутри) элемента [Action](/javascript/api/manifest/action) в объявлении элемента управления. Ниже показана базовая структура.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Изменение состояния программными средствами

Ниже приведены основные действия по изменению состояния "Включено" команды надстройки.

1. Создайте [объект RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) , который (1) задает команду и ее родительскую группу и вкладку по идентификаторам, объявленным в манифесте; и (2) указывает состояние включения или отключения команды.
2. Перенесите объект **RibbonUpdaterData** в метод [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)).

Ниже приведен простой пример. Обратите внимание, что myButton, OfficeAddinTab1 и CustomGroup111 копируются из манифеста.

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "OfficeAppTab1", 
                groups: [
                    {
                      id: "CustomGroup111",
                      controls: [
                        {
                            id: "MyButton", 
                            enabled: true
                        }
                      ]
                    }
                ]
            }
        ]
    });
}
```

Кроме того, мы предоставляем несколько интерфейсов (типов) для упрощения создания объекта **RibbonUpdateData**. Ниже приводится аналогичный пример в TypeScript, в котором используются эти типы.

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Вы можете `await` вызвать **requestUpdate(),** если родительская функция является асинхронной, но обратите внимание, что приложение Office управляет при обновлении состояния ленты. Метод **requestUpdate()** ставит запрос на обновление в очередь. Метод разрешает объект promise сразу после того, как он ставит запрос в очередь, а не при фактическом обновлении ленты.

## <a name="change-the-state-in-response-to-an-event"></a>Изменение состояния в ответ на событие

Обычно состояние ленты необходимо изменить, когда инициированное пользователем событие изменяет контекст надстройки.

Рассмотрим сценарий, в котором кнопка должна быть включена, только когда активирована диаграмма. Во-первых, задайте значение `false` для элемента [Enabled](/javascript/api/manifest/enabled) для кнопки в манифесте. Пример см. выше.

Во-вторых, назначьте обработчиков. Это обычно выполняется с помощью метода **Office.onReady**, как в приведенном ниже примере, где обработчики (созданные позднее) назначаются событиям **onActivated** и **onDeactivated** всех диаграмм на листе.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

В-третьих, определите обработчик `enableChartFormat`. Ниже приведен простой пример. Более надежный способ изменения состояния элемента управления см. в разделе [Рекомендация: проверка на наличие ошибок в состоянии элементов управления](#best-practice-test-for-control-status-errors) ниже.

```javascript
function enableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

В-четвертых, определите обработчик `disableChartFormat`. Он будет идентичен `enableChartFormat`, только для свойства объекта кнопки **enabled** будет задано значение `false`.

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Одновременное переключение видимости вкладки и состояния включенной кнопки

Метод **requestUpdate** также используется для переключения видимости настраиваемой контекстной вкладки. Дополнительные сведения об этом и примере кода см. в разделе ["Создание настраиваемых контекстных](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time) вкладок в надстройки Office".

## <a name="best-practice-test-for-control-status-errors"></a>Рекомендация: проверка на наличие ошибок в состоянии элементов управления

В некоторых случаях после вызова `requestUpdate` лента не обновляется, поэтому гиперсостояние элемента управления не изменяется. По этой причине рекомендуется отслеживать состояние элементов управления надстройки. Надстройка должна соответствовать следующим правилам.

1. При вызове `requestUpdate` в коде указывается предполагаемое состояние настраиваемых кнопок и элементов меню.
2. При щелчке пользовательского элемента управления первый код в обработчике проверяет, должна ли кнопка быть интерактивной. Если нет, код сообщит об ошибке или запишет ее в журнал и попробует еще раз установить для кнопок предполагаемое состояние.

В приведенном ниже примере показана функция, с помощью которой можно отключить кнопку и записать ее состояние. Обратите внимание, что `chartFormatButtonEnabled` — глобальная логическая переменная, которая инициализируется до того же значения, что и элемент [Enabled](/javascript/api/manifest/enabled) для кнопки в манифесте.

```javascript
function disableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

В приведенном ниже примере показано, как обработчик кнопки проверяет ее на наличие неправильного состояния. Обратите внимание, что `reportError` — это функция, которая отображает или записывает в журнал ошибку.

```javascript
function chartFormatButtonHandler() {
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
        var button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        var parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        var parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        var ribbonUpdater = {tabs: [parentTab]};
        Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
