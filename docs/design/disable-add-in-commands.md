---
title: Включение и отключение команд надстроек
description: Узнайте, как изменить состояние ("Включено" или "Отключено") настраиваемых кнопок ленты и элементов меню в веб-надстройке Office.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a2f2521b47a304fb358a33068bcccb425fd1a64
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511344"
---
# <a name="enable-and-disable-add-in-commands"></a>Включение и отключение команд надстроек

Если некоторые функции надстройки должны быть доступны только в определенном контексте, вы можете включить или отключить настраиваемые команды надстройки программными средствами. Например, функция, изменяющая заголовок таблицы, должна быть включена, только когда курсор находится в таблице.

Можно также указать, включена или отключена команда при Office клиентского приложения.

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Office только для приложений и платформ

API, описанные в этой статье, доступны только в Excel, PowerPoint и Word.

### <a name="test-for-platform-support-with-requirement-sets"></a>Тестирование поддержки платформ с использованием наборов обязательных элементов

Наборы требований — это именованные группы элементов API. Office надстройки используют наборы требований, указанные в манифесте, или используйте проверку времени запуска, чтобы определить, поддерживает ли сочетание приложений Office платформы API, необходимые надстройки. Дополнительные сведения см. [в Office версиях и наборах требований](../develop/office-versions-and-requirement-sets.md).

API enable/disable относятся к набору требований [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) .

> [!NOTE]
> Набор **требований RibbonApi 1.1** еще не поддерживается в манифесте, поэтому его нельзя указать в разделе **Требования манифеста** . Чтобы проверить поддержку, код должен вызвать `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`. Если и *только в том* случае, если этот вызов возвращается `true`, код может вызывать API- включить или отключить. Если вызов возвращается `isSetSupported` `false`, все настраиваемые команды надстройки включены все время. Необходимо разработать производственную надстройка и все инструкции в приложении, чтобы учитывать, как она будет работать, если набор требований **RibbonApi 1.1** не поддерживается. Дополнительные сведения `isSetSupported`и примеры использования см. в Office приложениях и требованиях [API, особенно](../develop/specify-office-hosts-and-api-requirements.md) проверках времени запуска для поддержки набора [методов и требований](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (Раздел [Укажите, какие Office](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) и платформы могут вмещая надстройки этой статьи, не применяется к Ленте 1.1.)

## <a name="shared-runtime-required"></a>Необходима общая среда выполнения

API и разметка манифеста надстройки, описанные в этой статье, требуют использования общей среды выполнения. Для этого необходимо предпринять следующие действия.

1. В элементе манифеста [Runtimes](../reference/manifest/runtimes.md) добавьте следующий дочерний элемент: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (Если в манифесте еще нет элемента **Runtimes** , создайте его в качестве первого ребенка в **элементе Host** в разделе **VersionOverrides** .)
2. В разделе [Resources.Urls](../reference/manifest/resources.md) манифеста добавьте следующий дочерний элемент:`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, где `{MyDomain}` домен надстройки и `{path-to-start-page}`путь к начальной странице надстройки; например: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. В зависимости от того, содержит ли надстройка области задач, файл функций или настраиваемую Excel, необходимо сделать один или несколько из следующих трех действий.

    - Если надстройка содержит области задач, установите `resid` атрибут [Действия](../reference/manifest/action.md).[ Элемент SourceLocation](../reference/manifest/sourcelocation.md) в точно такой же `resid` строке, как и для элемента **Runtime** на шаге 1; например, `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если надстройка содержит Excel настраиваемую функцию, `resid` установите атрибут [Страницы](../reference/manifest/page.md).[ Элемент SourceLocation](../reference/manifest/sourcelocation.md) точно такой же строки`resid`, как и для элемента **Runtime** на шаге 1; например, `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Если надстройка содержит файл функции, `resid` установите атрибут элемента [FunctionFile](../reference/manifest/functionfile.md) `resid` точно такой же строкой, как вы использовали для элемента **Runtime** на шаге 1; например, `Contoso.SharedRuntime.Url`. Элемент должен выглядеть следующим образом:`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

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
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Изменение состояния программными средствами

Ниже приведены основные действия по изменению состояния "Включено" команды надстройки.

1. Создание [объекта RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) , который (1) указывает команду и ее родительскую группу и вкладку по своим ID-адресам, как указано в манифесте; и (2) указывает состояние включенной или отключенной команды.
2. Перенесите объект **RibbonUpdaterData** в метод [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)).

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

Вы можете `await` вызвать **requestUpdate()** если родительская функция асинхронна, но обратите внимание, что Office приложение управляет при обновлении состояния ленты. Метод **requestUpdate()** ставит запрос на обновление в очередь. Метод разрешит объект promise сразу после очереди запроса, а не после обновления ленты.

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

Метод **requestUpdate** также используется для настройки видимости настраиваемой контекстной вкладки. Дополнительные сведения об этом и примере кода см. в материале [Create custom contextual tabs in Office надстроек](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).

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
