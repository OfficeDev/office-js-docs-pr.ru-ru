---
title: Включение и отключение команд надстроек
description: Узнайте, как изменить состояние ("Включено" или "Отключено") настраиваемых кнопок ленты и элементов меню в веб-надстройке Office.
ms.date: 04/11/2020
localization_priority: Priority
ms.openlocfilehash: a0436a07ef5c7ec64ad391747da69061e1a7b0f0
ms.sourcegitcommit: 231e23d72e04e0536480d6b16df95113f1eff738
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2020
ms.locfileid: "43238229"
---
# <a name="enable-and-disable-add-in-commands-preview"></a><span data-ttu-id="4ff4b-103">Включение и отключение команд надстроек (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="4ff4b-103">Enable and Disable Add-in Commands (preview)</span></span>

<span data-ttu-id="4ff4b-104">Если некоторые функции надстройки должны быть доступны только в определенном контексте, вы можете включить или отключить настраиваемые команды надстройки программными средствами.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="4ff4b-105">Например, функция, изменяющая заголовок таблицы, должна быть включена, только когда курсор находится в таблице.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="4ff4b-106">Также можно указать, будет ли команда включена или отключена при открытии ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-106">You can also specify whether the command is enabled or disabled when the Office host application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="4ff4b-107">В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="4ff4b-108">Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> [<span data-ttu-id="4ff4b-109">Основные концепции команд надстроек</span><span class="sxs-lookup"><span data-stu-id="4ff4b-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="preview-status"></a><span data-ttu-id="4ff4b-110">Состояние предварительной версии</span><span class="sxs-lookup"><span data-stu-id="4ff4b-110">Preview status</span></span>

<span data-ttu-id="4ff4b-111">API, описанные в этой статье, находятся в предварительной версии и в настоящее время доступны только в Excel.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-111">The APIs described in this article are in preview and are currently only available in Excel.</span></span>

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="rules-and-gotchas"></a><span data-ttu-id="4ff4b-112">Правила и подсказки</span><span class="sxs-lookup"><span data-stu-id="4ff4b-112">Rules and gotchas</span></span>

### <a name="single-line-ribbon-in-office-on-the-web"></a><span data-ttu-id="4ff4b-113">Однострочная лента в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="4ff4b-113">Single-line ribbon in Office on the web</span></span>

<span data-ttu-id="4ff4b-114">В Office в Интернете API и разметка манифеста, описанные в этой статье, применимы только к однострочной ленте.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-114">In Office on the web, the APIs and manifest markup described in this article only affect the single-line ribbon.</span></span> <span data-ttu-id="4ff4b-115">Они не оказывают влияния на многострочную ленту.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-115">They have no effect on the multiline ribbon.</span></span> <span data-ttu-id="4ff4b-116">Они затрагивают обе ленты в классических приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-116">They affect both ribbons for desktop Office.</span></span> <span data-ttu-id="4ff4b-117">Дополнительные сведения об обеих лентах см. в статье [Использование упрощенной ленты](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span><span class="sxs-lookup"><span data-stu-id="4ff4b-117">For more information about the two ribbons, see [Use the simplified ribbon](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span></span>

### <a name="shared-runtime-required"></a><span data-ttu-id="4ff4b-118">Необходима общая среда выполнения</span><span class="sxs-lookup"><span data-stu-id="4ff4b-118">Shared runtime required</span></span>

<span data-ttu-id="4ff4b-119">API и разметка манифеста надстройки, описанные в этой статье, требуют использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-119">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="4ff4b-120">Для этого выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-120">To do this take the following steps.</span></span>

1. <span data-ttu-id="4ff4b-121">В элементе манифеста [Runtimes](../reference/manifest/runtimes.md) добавьте следующий дочерний элемент: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-121">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="4ff4b-122">(Если в манифесте еще нет элемента `<Runtimes>`, создайте его в качестве первого дочернего элемента `<Host>` в разделе `VersionOverrides`.)</span><span class="sxs-lookup"><span data-stu-id="4ff4b-122">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="4ff4b-123">В разделе [Resources.Urls](../reference/manifest/resources.md) манифеста добавьте следующий дочерний элемент:`<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, где `{MyDomain}` домен надстройки и `{path-to-start-page}`путь к начальной странице надстройки; например: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-123">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="4ff4b-124">В зависимости от того, есть ли в вашей надстройке область задач, файл функций или настраиваемая функция Excel, необходимо выполнить одно или несколько из описанных ниже трех действий.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-124">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="4ff4b-125">Если в надстройке есть область задач, задайте атрибут `resid` элемента [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md)точно так же, как в `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-125">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="4ff4b-126">Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-126">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="4ff4b-127">Если в настройке есть настраиваемая функция Excel, установите атрибут `resid` элемента [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-127">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="4ff4b-128">Элемент должен выглядеть следующим образом:`<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-128">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="4ff4b-129">Если в настройке есть файл функций, установите атрибут `resid` элемента [FunctionFile](../reference/manifest/functionfile.md) точно так же, как для `resid` элемента `<Runtime>` на шаге 1; например `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-129">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="4ff4b-130">Элемент должен выглядеть следующим образом:`<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-130">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="4ff4b-131">Установка состояния "Отключено" по умолчанию</span><span class="sxs-lookup"><span data-stu-id="4ff4b-131">Set the default state to disabled</span></span>

<span data-ttu-id="4ff4b-132">По умолчанию при запуске приложения Office любая команда надстройки включается.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-132">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="4ff4b-133">Если вы хотите, чтобы при запуске приложения Office настраиваемая кнопка или элемент меню были отключены, укажите это в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-133">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="4ff4b-134">Просто добавьте элемент [Enabled](../reference/manifest/enabled.md) (со значением`false`) сразу *под* (не внутри) элемента [Action](../reference/manifest/action.md) в объявлении элемента управления.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-134">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="4ff4b-135">Ниже показана базовая структура.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-135">The following shows the basic structure:</span></span>

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

## <a name="change-the-state-programmatically"></a><span data-ttu-id="4ff4b-136">Изменение состояния программными средствами</span><span class="sxs-lookup"><span data-stu-id="4ff4b-136">Change the state programmatically</span></span>

<span data-ttu-id="4ff4b-137">Ниже приведены основные действия по изменению состояния "Включено" команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-137">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="4ff4b-138">Создайте объект [RibbonUpdaterData](/javascript/api/office-runtime/officeruntime.ribbonupdaterdata), в котором (1) указаны идентификаторы команды и ее родительской вкладки в соответствии с манифестом и (2) указано состояние команды ("Включено" или "Отключено").</span><span class="sxs-lookup"><span data-stu-id="4ff4b-138">Create a [RibbonUpdaterData](/javascript/api/office-runtime/officeruntime.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="4ff4b-139">Перенесите объект **RibbonUpdaterData** в метод [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-).</span><span class="sxs-lookup"><span data-stu-id="4ff4b-139">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) method.</span></span>

<span data-ttu-id="4ff4b-140">Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-140">The following is a simple example.</span></span> <span data-ttu-id="4ff4b-141">Обратите внимание, что "MyButton" и "OfficeAddinTab1" скопированы из манифеста.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-141">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

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

<span data-ttu-id="4ff4b-142">Кроме того, мы предоставляем несколько интерфейсов (типов) для упрощения создания объекта **RibbonUpdateData**.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-142">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="4ff4b-143">Ниже приводится аналогичный пример в TypeScript, в котором используются эти типы.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-143">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="4ff4b-144">Office определяет время обновления состояния ленты.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-144">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="4ff4b-145">Метод **requestUpdate()** ставит запрос на обновление в очередь.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-145">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="4ff4b-146">Этот метод устранит объект Promise, как только он поставит запрос в очередь, а не при обновлении ленты.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-146">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="4ff4b-147">Изменение состояния в ответ на событие</span><span class="sxs-lookup"><span data-stu-id="4ff4b-147">Change the state in response to an event</span></span>

<span data-ttu-id="4ff4b-148">Обычно состояние ленты необходимо изменить, когда инициированное пользователем событие изменяет контекст надстройки.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-148">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="4ff4b-149">Рассмотрим сценарий, в котором кнопка должна быть включена, только когда активирована диаграмма.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-149">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="4ff4b-150">Во-первых, задайте значение `false` для элемента [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-150">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="4ff4b-151">Пример см. выше.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-151">See above for an example.</span></span>

<span data-ttu-id="4ff4b-152">Во-вторых, назначьте обработчиков.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-152">Second, assign handlers.</span></span> <span data-ttu-id="4ff4b-153">Это обычно выполняется с помощью метода **Office.onReady**, как в приведенном ниже примере, где обработчики (созданные позднее) назначаются событиям **onActivated** и **onDeactivated** всех диаграмм на листе.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-153">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="4ff4b-154">В-третьих, определите обработчик `enableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-154">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="4ff4b-155">Ниже приведен простой пример. Более надежный способ изменения состояния элемента управления см. в разделе **Рекомендация: проверка на наличие ошибок в состоянии элементов управления** ниже.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-155">The following is a simple example, but see **Best practice: Test for control status errors** below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="4ff4b-156">В-четвертых, определите обработчик `disableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-156">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="4ff4b-157">Он будет идентичен `enableChartFormat`, только для свойства объекта кнопки **enabled** будет задано значение `false`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-157">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="4ff4b-158">Рекомендация: проверка на наличие ошибок в состоянии элементов управления</span><span class="sxs-lookup"><span data-stu-id="4ff4b-158">Best practice: Test for control status errors</span></span>

<span data-ttu-id="4ff4b-159">В некоторых случаях после вызова `requestUpdate` лента не обновляется, поэтому гиперсостояние элемента управления не изменяется.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-159">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="4ff4b-160">По этой причине рекомендуется отслеживать состояние элементов управления надстройки.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-160">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="4ff4b-161">Надстройка должна соответствовать приведенным ниже требованиям.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-161">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="4ff4b-162">При вызове `requestUpdate` в коде указывается предполагаемое состояние настраиваемых кнопок и элементов меню.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-162">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="4ff4b-163">При щелчке пользовательского элемента управления первый код в обработчике проверяет, должна ли кнопка быть интерактивной.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-163">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="4ff4b-164">Если нет, код сообщит об ошибке или запишет ее в журнал и попробует еще раз установить для кнопок предполагаемое состояние.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-164">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="4ff4b-165">В приведенном ниже примере показана функция, с помощью которой можно отключить кнопку и записать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-165">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="4ff4b-166">Обратите внимание, что `chartFormatButtonEnabled` — глобальная логическая переменная, которая инициализируется до того же значения, что и элемент [Enabled](../reference/manifest/enabled.md) для кнопки в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-166">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="4ff4b-167">В приведенном ниже примере показано, как обработчик кнопки проверяет ее на наличие неправильного состояния.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-167">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="4ff4b-168">Обратите внимание, что `reportError` — это функция, которая отображает или записывает в журнал ошибку.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-168">Note that `reportError` is a function that shows or logs an error.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="4ff4b-169">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="4ff4b-169">Error handling</span></span>

<span data-ttu-id="4ff4b-170">В некоторых случаях Office не может обновить ленту и возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-170">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="4ff4b-171">Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-171">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="4ff4b-172">Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-172">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="4ff4b-173">Ниже приведен пример обработки этой ошибки.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-173">The following is an example of how to handle this error.</span></span> <span data-ttu-id="4ff4b-174">В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.</span><span class="sxs-lookup"><span data-stu-id="4ff4b-174">In this case, the `reportError` method displays the error to the user.</span></span>

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

## <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="4ff4b-175">Тестирование поддержки платформ с использованием наборов обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="4ff4b-175">Test for platform support with requirement sets</span></span>

<span data-ttu-id="4ff4b-p122">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="4ff4b-p122">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="4ff4b-179">Для включения или отключения API требуется поддержка следующего набора обязательных элементов:</span><span class="sxs-lookup"><span data-stu-id="4ff4b-179">The enable/disable APIs require support of the following requirement set:</span></span>

- [<span data-ttu-id="4ff4b-180">AddinCommands 1.1</span><span class="sxs-lookup"><span data-stu-id="4ff4b-180">AddinCommands 1.1</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
