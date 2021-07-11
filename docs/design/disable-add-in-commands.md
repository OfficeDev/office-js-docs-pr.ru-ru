---
title: Включение и отключение команд надстроек
description: Узнайте, как изменить состояние ("Включено" или "Отключено") настраиваемых кнопок ленты и элементов меню в веб-надстройке Office.
ms.date: 04/30/2021
localization_priority: Normal
ms.openlocfilehash: 2a2816990a7f21a4238a9f8332537bf904fa4cb2
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349226"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="5ccec-103">Включение и отключение команд надстроек</span><span class="sxs-lookup"><span data-stu-id="5ccec-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="5ccec-104">Если некоторые функции надстройки должны быть доступны только в определенном контексте, вы можете включить или отключить настраиваемые команды надстройки программными средствами.</span><span class="sxs-lookup"><span data-stu-id="5ccec-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="5ccec-105">Например, функция, изменяющая заголовок таблицы, должна быть включена, только когда курсор находится в таблице.</span><span class="sxs-lookup"><span data-stu-id="5ccec-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="5ccec-106">Можно также указать, включена или отключена команда при Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="5ccec-106">You can also specify whether the command is enabled or disabled when the Office client application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="5ccec-107">В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией.</span><span class="sxs-lookup"><span data-stu-id="5ccec-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="5ccec-108">Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.</span><span class="sxs-lookup"><span data-stu-id="5ccec-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="5ccec-109">Основные концепции команд надстроек</span><span class="sxs-lookup"><span data-stu-id="5ccec-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a><span data-ttu-id="5ccec-110">Office только для приложений и платформ</span><span class="sxs-lookup"><span data-stu-id="5ccec-110">Office application and platform support only</span></span>

<span data-ttu-id="5ccec-111">API, описанные в этой статье, доступны только в Excel на всех платформах и в PowerPoint в Интернете.</span><span class="sxs-lookup"><span data-stu-id="5ccec-111">The APIs described in this article are only available in Excel on all platforms and in PowerPoint on the web.</span></span>

### <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="5ccec-112">Тестирование поддержки платформ с использованием наборов обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="5ccec-112">Test for platform support with requirement sets</span></span>

<span data-ttu-id="5ccec-113">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="5ccec-113">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="5ccec-114">Office Надстройки используют наборы требований, указанные в манифесте, или используйте проверку времени запуска, чтобы определить, поддерживает ли приложение Office и платформа API, которые необходимы надстройки.</span><span class="sxs-lookup"><span data-stu-id="5ccec-114">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application and platform combination supports APIs that an add-in needs.</span></span> <span data-ttu-id="5ccec-115">Дополнительные сведения см. [в Office версиях и наборах требований.](../develop/office-versions-and-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="5ccec-115">For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="5ccec-116">API enable/disable относятся к набору требований [RibbonApi 1.1.](../reference/requirement-sets/ribbon-api-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="5ccec-116">The enable/disable APIs belong to the [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) requirement set.</span></span>

> [!NOTE]
> <span data-ttu-id="5ccec-117">Набор **требований RibbonApi 1.1** еще не поддерживается в манифесте, поэтому его нельзя указать в разделе `<Requirements>` манифест.</span><span class="sxs-lookup"><span data-stu-id="5ccec-117">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span> <span data-ttu-id="5ccec-118">Чтобы проверить поддержку, код должен вызвать `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` .</span><span class="sxs-lookup"><span data-stu-id="5ccec-118">To test for support, your code should call `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`.</span></span> <span data-ttu-id="5ccec-119">Если и *только в том* случае, если этот вызов возвращается, код может вызывать `true` API- включить или отключить.</span><span class="sxs-lookup"><span data-stu-id="5ccec-119">If, *and only if*, that call returns `true`, your code can call the enable/disable APIs.</span></span> <span data-ttu-id="5ccec-120">Если вызов `isSetSupported` возвращается, все настраиваемые команды надстройки включены `false` все время.</span><span class="sxs-lookup"><span data-stu-id="5ccec-120">If the call of `isSetSupported` returns `false`, then all custom add-in commands are enabled all of the time.</span></span> <span data-ttu-id="5ccec-121">Необходимо разработать производственную надстройка и все инструкции в приложении, чтобы учесть, как она будет работать, если набор требований **RibbonApi 1.1** не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5ccec-121">You must design your production add-in, and any in-app instructions, to take account of how it will work when the **RibbonApi 1.1** requirement set is not supported.</span></span> <span data-ttu-id="5ccec-122">Дополнительные сведения и примеры использования см. в Office приложениях и требованиях API, особенно используйте проверки времени запуска в `isSetSupported` [коде JavaScript.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) [](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="5ccec-122">For more information and examples of using `isSetSupported`, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md), especially [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="5ccec-123">(В разделе [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)</span><span class="sxs-lookup"><span data-stu-id="5ccec-123">(The section [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)</span></span>

## <a name="shared-runtime-required"></a><span data-ttu-id="5ccec-124">Необходима общая среда выполнения</span><span class="sxs-lookup"><span data-stu-id="5ccec-124">Shared runtime required</span></span>

<span data-ttu-id="5ccec-125">API и разметка манифеста надстройки, описанные в этой статье, требуют использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="5ccec-125">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="5ccec-126">Для этого выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="5ccec-126">To do this take the following steps.</span></span>

1. <span data-ttu-id="5ccec-127">В элементе манифеста [Runtimes](../reference/manifest/runtimes.md) добавьте следующий дочерний элемент: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-127">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="5ccec-128">(Если в манифесте еще нет элемента `<Runtimes>`, создайте его в качестве первого дочернего элемента `<Host>` в разделе `VersionOverrides`.)</span><span class="sxs-lookup"><span data-stu-id="5ccec-128">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="5ccec-129">В разделе [Resources.Urls](../reference/manifest/resources.md) манифеста добавьте следующий дочерний элемент:`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, где `{MyDomain}` домен надстройки и `{path-to-start-page}`путь к начальной странице надстройки; например: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-129">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="5ccec-130">В зависимости от того, содержит ли надстройка области задач, файл функций или настраиваемую Excel, необходимо сделать один или несколько из следующих трех действий.</span><span class="sxs-lookup"><span data-stu-id="5ccec-130">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps.</span></span>

    - <span data-ttu-id="5ccec-131">Если в надстройке есть область задач, задайте атрибут `resid` элемента [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md)точно так же, как в `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-131">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="5ccec-132">Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-132">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="5ccec-133">Если в настройке есть настраиваемая функция Excel, установите атрибут `resid` элемента [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-133">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="5ccec-134">Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-134">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="5ccec-135">Если в настройке есть файл функций, установите атрибут `resid` элемента [FunctionFile](../reference/manifest/functionfile.md) точно так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-135">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="5ccec-136">Элемент должен выглядеть следующим образом:`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-136">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="5ccec-137">Установка состояния "Отключено" по умолчанию</span><span class="sxs-lookup"><span data-stu-id="5ccec-137">Set the default state to disabled</span></span>

<span data-ttu-id="5ccec-138">По умолчанию при запуске приложения Office любая команда надстройки включается.</span><span class="sxs-lookup"><span data-stu-id="5ccec-138">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="5ccec-139">Если вы хотите, чтобы при запуске приложения Office настраиваемая кнопка или элемент меню были отключены, укажите это в манифесте.</span><span class="sxs-lookup"><span data-stu-id="5ccec-139">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="5ccec-140">Просто добавьте элемент [Enabled](../reference/manifest/enabled.md) (со значением`false`) сразу *под* (не внутри) элемента [Action](../reference/manifest/action.md) в объявлении элемента управления.</span><span class="sxs-lookup"><span data-stu-id="5ccec-140">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="5ccec-141">Ниже показана базовая структура.</span><span class="sxs-lookup"><span data-stu-id="5ccec-141">The following shows the basic structure.</span></span>

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

## <a name="change-the-state-programmatically"></a><span data-ttu-id="5ccec-142">Изменение состояния программными средствами</span><span class="sxs-lookup"><span data-stu-id="5ccec-142">Change the state programmatically</span></span>

<span data-ttu-id="5ccec-143">Ниже приведены основные действия по изменению состояния "Включено" команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="5ccec-143">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="5ccec-144">Создание [объекта RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) указывает команду и ее родительскую группу и вкладку по своим ID-адресам, как указано в манифесте; и (2) указывает состояние включенной или отключенной команды.</span><span class="sxs-lookup"><span data-stu-id="5ccec-144">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent group and tab, by their IDs as declared in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="5ccec-145">Перенесите объект **RibbonUpdaterData** в метод [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-).</span><span class="sxs-lookup"><span data-stu-id="5ccec-145">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method.</span></span>

<span data-ttu-id="5ccec-146">Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="5ccec-146">The following is a simple example.</span></span> <span data-ttu-id="5ccec-147">Обратите внимание, что из манифеста копируется "MyButton", "OfficeAddinTab1" и "CustomGroup111".</span><span class="sxs-lookup"><span data-stu-id="5ccec-147">Note that "MyButton", "OfficeAddinTab1", and "CustomGroup111" are copied from the manifest.</span></span>

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

<span data-ttu-id="5ccec-148">Кроме того, мы предоставляем несколько интерфейсов (типов) для упрощения создания объекта **RibbonUpdateData**.</span><span class="sxs-lookup"><span data-stu-id="5ccec-148">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="5ccec-149">Ниже приводится аналогичный пример в TypeScript, в котором используются эти типы.</span><span class="sxs-lookup"><span data-stu-id="5ccec-149">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="5ccec-150">Вы можете вызвать `await` **requestUpdate()** если родительская функция асинхронна, но обратите внимание, что Office управления приложением при обновлении состояния ленты.</span><span class="sxs-lookup"><span data-stu-id="5ccec-150">You can `await` the call of **requestUpdate()** if the parent function is asynchronous, but note that the Office application controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="5ccec-151">Метод **requestUpdate()** ставит запрос на обновление в очередь.</span><span class="sxs-lookup"><span data-stu-id="5ccec-151">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="5ccec-152">Метод разрешит объект promise сразу после очереди запроса, а не после обновления ленты.</span><span class="sxs-lookup"><span data-stu-id="5ccec-152">The method will resolve the promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="5ccec-153">Изменение состояния в ответ на событие</span><span class="sxs-lookup"><span data-stu-id="5ccec-153">Change the state in response to an event</span></span>

<span data-ttu-id="5ccec-154">Обычно состояние ленты необходимо изменить, когда инициированное пользователем событие изменяет контекст надстройки.</span><span class="sxs-lookup"><span data-stu-id="5ccec-154">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="5ccec-155">Рассмотрим сценарий, в котором кнопка должна быть включена, только когда активирована диаграмма.</span><span class="sxs-lookup"><span data-stu-id="5ccec-155">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="5ccec-156">Во-первых, задайте значение `false` для элемента [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте.</span><span class="sxs-lookup"><span data-stu-id="5ccec-156">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="5ccec-157">Пример см. выше.</span><span class="sxs-lookup"><span data-stu-id="5ccec-157">See above for an example.</span></span>

<span data-ttu-id="5ccec-158">Во-вторых, назначьте обработчиков.</span><span class="sxs-lookup"><span data-stu-id="5ccec-158">Second, assign handlers.</span></span> <span data-ttu-id="5ccec-159">Это обычно выполняется с помощью метода **Office.onReady**, как в приведенном ниже примере, где обработчики (созданные позднее) назначаются событиям **onActivated** и **onDeactivated** всех диаграмм на листе.</span><span class="sxs-lookup"><span data-stu-id="5ccec-159">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="5ccec-160">В-третьих, определите обработчик `enableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-160">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="5ccec-161">Ниже приведен простой пример. Более надежный способ изменения состояния элемента управления см. в разделе [Рекомендация: проверка на наличие ошибок в состоянии элементов управления](#best-practice-test-for-control-status-errors) ниже.</span><span class="sxs-lookup"><span data-stu-id="5ccec-161">The following is a simple example, but see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors) below for a more robust way of changing a control's status.</span></span>

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

<span data-ttu-id="5ccec-162">В-четвертых, определите обработчик `disableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-162">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="5ccec-163">Он будет идентичен `enableChartFormat`, только для свойства объекта кнопки **enabled** будет задано значение `false`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-163">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="5ccec-164">Обзор вкладок и состояние включенной кнопки одновременно</span><span class="sxs-lookup"><span data-stu-id="5ccec-164">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="5ccec-165">Метод **requestUpdate** также используется для настройки видимости настраиваемой контекстной вкладки. Дополнительные сведения об этом и примере кода см. в материале [Create custom contextual tabs in Office надстроек.](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)</span><span class="sxs-lookup"><span data-stu-id="5ccec-165">The **requestUpdate** method is also used to toggle the visibility of a custom contextual tab. For details about this and example code, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="5ccec-166">Рекомендация: проверка на наличие ошибок в состоянии элементов управления</span><span class="sxs-lookup"><span data-stu-id="5ccec-166">Best practice: Test for control status errors</span></span>

<span data-ttu-id="5ccec-167">В некоторых случаях после вызова `requestUpdate` лента не обновляется, поэтому гиперсостояние элемента управления не изменяется.</span><span class="sxs-lookup"><span data-stu-id="5ccec-167">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="5ccec-168">По этой причине рекомендуется отслеживать состояние элементов управления надстройки.</span><span class="sxs-lookup"><span data-stu-id="5ccec-168">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="5ccec-169">Надстройка должна соответствовать приведенным ниже требованиям.</span><span class="sxs-lookup"><span data-stu-id="5ccec-169">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="5ccec-170">При вызове `requestUpdate` в коде указывается предполагаемое состояние настраиваемых кнопок и элементов меню.</span><span class="sxs-lookup"><span data-stu-id="5ccec-170">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="5ccec-171">При щелчке пользовательского элемента управления первый код в обработчике проверяет, должна ли кнопка быть интерактивной.</span><span class="sxs-lookup"><span data-stu-id="5ccec-171">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="5ccec-172">Если нет, код сообщит об ошибке или запишет ее в журнал и попробует еще раз установить для кнопок предполагаемое состояние.</span><span class="sxs-lookup"><span data-stu-id="5ccec-172">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="5ccec-173">В приведенном ниже примере показана функция, с помощью которой можно отключить кнопку и записать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="5ccec-173">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="5ccec-174">Обратите внимание, что `chartFormatButtonEnabled` — глобальная логическая переменная, которая инициализируется до того же значения, что и элемент [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте.</span><span class="sxs-lookup"><span data-stu-id="5ccec-174">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

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

<span data-ttu-id="5ccec-175">В приведенном ниже примере показано, как обработчик кнопки проверяет ее на наличие неправильного состояния.</span><span class="sxs-lookup"><span data-stu-id="5ccec-175">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="5ccec-176">Обратите внимание, что `reportError` — это функция, которая отображает или записывает в журнал ошибку.</span><span class="sxs-lookup"><span data-stu-id="5ccec-176">Note that `reportError` is a function that shows or logs an error.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="5ccec-177">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="5ccec-177">Error handling</span></span>

<span data-ttu-id="5ccec-178">В некоторых случаях Office не может обновить ленту и возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="5ccec-178">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="5ccec-179">Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть.</span><span class="sxs-lookup"><span data-stu-id="5ccec-179">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="5ccec-180">Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="5ccec-180">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="5ccec-181">Ниже приведен пример обработки этой ошибки.</span><span class="sxs-lookup"><span data-stu-id="5ccec-181">The following is an example of how to handle this error.</span></span> <span data-ttu-id="5ccec-182">В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.</span><span class="sxs-lookup"><span data-stu-id="5ccec-182">In this case, the `reportError` method displays the error to the user.</span></span>

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
