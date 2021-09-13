---
title: Включение и отключение команд надстроек
description: Узнайте, как изменить состояние ("Включено" или "Отключено") настраиваемых кнопок ленты и элементов меню в веб-надстройке Office.
ms.date: 07/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: f69f293292eb3ce04e366f740cbd92ab9fe13617
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151058"
---
# <a name="enable-and-disable-add-in-commands"></a>Включение и отключение команд надстроек

Если некоторые функции надстройки должны быть доступны только в определенном контексте, вы можете включить или отключить настраиваемые команды надстройки программными средствами. Например, функция, изменяющая заголовок таблицы, должна быть включена, только когда курсор находится в таблице.

Можно также указать, включена или отключена команда при Office клиентского приложения.

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

[!INCLUDE [Animation of contextual tabs and enabling buttons](../includes/animation-contextual-tabs-enable-button.md)]

## <a name="office-application-and-platform-support-only"></a>Office только для приложений и платформ

API, описанные в этой статье, доступны только в Excel на всех платформах и в PowerPoint в Интернете.

### <a name="test-for-platform-support-with-requirement-sets"></a>Тестирование поддержки платформ с использованием наборов обязательных элементов

Наборы требований — это именованные группы элементов API. Office Надстройки используют наборы требований, указанные в манифесте, или используйте проверку времени запуска, чтобы определить, поддерживает ли приложение Office и платформа API, которые необходимы надстройки. Дополнительные сведения см. [в Office версиях и наборах требований.](../develop/office-versions-and-requirement-sets.md)

API enable/disable относятся к набору требований [RibbonApi 1.1.](../reference/requirement-sets/ribbon-api-requirement-sets.md)

> [!NOTE]
> Набор **требований RibbonApi 1.1** еще не поддерживается в манифесте, поэтому его нельзя указать в разделе `<Requirements>` манифест. Чтобы проверить поддержку, код должен вызвать `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` . Если и *только в том* случае, если этот вызов возвращается, код может вызывать `true` API- включить или отключить. Если вызов `isSetSupported` возвращается, все настраиваемые команды надстройки включены `false` все время. Необходимо разработать производственную надстройка и все инструкции в приложении, чтобы учесть, как она будет работать, если набор требований **RibbonApi 1.1** не поддерживается. Дополнительные сведения и примеры использования см. в Office приложениях и требованиях API, особенно используйте проверки времени запуска в `isSetSupported` [коде JavaScript.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) [](../develop/specify-office-hosts-and-api-requirements.md) (В разделе [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)

## <a name="shared-runtime-required"></a>Необходима общая среда выполнения

API и разметка манифеста надстройки, описанные в этой статье, требуют использования общей среды выполнения. Для этого необходимо предпринять следующие действия.

1. В элементе манифеста [Runtimes](../reference/manifest/runtimes.md) добавьте следующий дочерний элемент: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (Если в манифесте еще нет элемента `<Runtimes>`, создайте его в качестве первого дочернего элемента `<Host>` в разделе `VersionOverrides`.)
2. В разделе [Resources.Urls](../reference/manifest/resources.md) манифеста добавьте следующий дочерний элемент:`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, где `{MyDomain}` домен надстройки и `{path-to-start-page}`путь к начальной странице надстройки; например: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. В зависимости от того, содержит ли надстройка области задач, файл функций или настраиваемую Excel, необходимо сделать один или несколько из следующих трех действий.

    - Если в надстройке есть область задач, задайте атрибут `resid` элемента [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md)точно так же, как в `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если в настройке есть настраиваемая функция Excel, установите атрибут `resid` элемента [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если в настройке есть файл функций, установите атрибут `resid` элемента [FunctionFile](../reference/manifest/functionfile.md) точно так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Установка состояния "Отключено" по умолчанию

По умолчанию при запуске приложения Office любая команда надстройки включается. Если вы хотите, чтобы при запуске приложения Office настраиваемая кнопка или элемент меню были отключены, укажите это в манифесте. Просто добавьте элемент [Enabled](../reference/manifest/enabled.md) (со значением`false`) сразу *под* (не внутри) элемента [Action](../reference/manifest/action.md) в объявлении элемента управления. Ниже показана базовая структура.

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
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Изменение состояния программными средствами

Ниже приведены основные действия по изменению состояния "Включено" команды надстройки.

1. Создание [объекта RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) указывает команду и ее родительскую группу и вкладку по своим ID-адресам, как указано в манифесте; и (2) указывает состояние включенной или отключенной команды.
2. Перенесите объект **RibbonUpdaterData** в метод [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestUpdate_input_).

Ниже приведен простой пример. Обратите внимание, что из манифеста копируется "MyButton", "OfficeAddinTab1" и "CustomGroup111".

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

Вы можете вызвать `await` **requestUpdate()** если родительская функция асинхронна, но обратите внимание, что Office управления приложением при обновлении состояния ленты. Метод **requestUpdate()** ставит запрос на обновление в очередь. Метод разрешит объект promise сразу после очереди запроса, а не после обновления ленты.

## <a name="change-the-state-in-response-to-an-event"></a>Изменение состояния в ответ на событие

Обычно состояние ленты необходимо изменить, когда инициированное пользователем событие изменяет контекст надстройки.

Рассмотрим сценарий, в котором кнопка должна быть включена, только когда активирована диаграмма. Во-первых, задайте значение `false` для элемента [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте. Пример см. выше.

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

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Обзор вкладок и состояние включенной кнопки одновременно

Метод **requestUpdate** также используется для настройки видимости настраиваемой контекстной вкладки. Дополнительные сведения об этом и примере кода см. в материале [Create custom contextual tabs in Office надстроек.](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)

## <a name="best-practice-test-for-control-status-errors"></a>Рекомендация: проверка на наличие ошибок в состоянии элементов управления

В некоторых случаях после вызова `requestUpdate` лента не обновляется, поэтому гиперсостояние элемента управления не изменяется. По этой причине рекомендуется отслеживать состояние элементов управления надстройки. Надстройка должна соответствовать следующим правилам.

1. При вызове `requestUpdate` в коде указывается предполагаемое состояние настраиваемых кнопок и элементов меню.
2. При щелчке пользовательского элемента управления первый код в обработчике проверяет, должна ли кнопка быть интерактивной. Если нет, код сообщит об ошибке или запишет ее в журнал и попробует еще раз установить для кнопок предполагаемое состояние.

В приведенном ниже примере показана функция, с помощью которой можно отключить кнопку и записать ее состояние. Обратите внимание, что `chartFormatButtonEnabled` — глобальная логическая переменная, которая инициализируется до того же значения, что и элемент [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте.

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
