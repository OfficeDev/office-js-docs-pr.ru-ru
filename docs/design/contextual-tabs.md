---
title: Создание настраиваемой контекстной вкладки в надстройки Office
description: Узнайте, как добавлять настраиваемые контекстные вкладки в надстройку Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 67588e04d6ea95bc581c51e274c8135cfa5afd50
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173922"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="047bb-103">Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="047bb-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="047bb-104">Контекстная вкладка — это скрытая вкладка на ленте Office, которая отображается в строке вкладки, когда в документе Office происходит определенное событие.</span><span class="sxs-lookup"><span data-stu-id="047bb-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="047bb-105">Например, **вкладка "Конструктор таблицы",** которая отображается на ленте Excel при выборе таблицы.</span><span class="sxs-lookup"><span data-stu-id="047bb-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="047bb-106">Вы можете включить настраиваемые контекстные вкладки в надстройку Office и указать, когда они будут видимыми или скрытыми, создав обработчики событий, которые меняют видимость.</span><span class="sxs-lookup"><span data-stu-id="047bb-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="047bb-107">(Однако настраиваемые контекстные вкладки не реагируют на изменения фокуса.)</span><span class="sxs-lookup"><span data-stu-id="047bb-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="047bb-108">В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией.</span><span class="sxs-lookup"><span data-stu-id="047bb-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="047bb-109">Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.</span><span class="sxs-lookup"><span data-stu-id="047bb-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="047bb-110">Основные концепции команд надстроек</span><span class="sxs-lookup"><span data-stu-id="047bb-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="047bb-111">Настраиваемые контекстные вкладки находятся в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="047bb-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="047bb-112">Поэкспериментируйте с ними в среде разработки или тестирования, но не добавляйте их в производственную надстройки.</span><span class="sxs-lookup"><span data-stu-id="047bb-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="047bb-113">Настраиваемые контекстные вкладки в настоящее время поддерживаются только в Excel и только на этих платформах и сборках:</span><span class="sxs-lookup"><span data-stu-id="047bb-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="047bb-114">Excel для Windows (только Microsoft 365, а не бессрочная лицензия): версия 2011 (сборка 13426.20274).</span><span class="sxs-lookup"><span data-stu-id="047bb-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="047bb-115">Возможно, ваша подписка на Microsoft 365 должна быть на канале [Current Channel (предварительная версия),](https://insider.office.com/join/windows) который ранее назывался Monthly Channel (Targeted)" или "Insider Slow".</span><span class="sxs-lookup"><span data-stu-id="047bb-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="047bb-116">Настраиваемые контекстные вкладки работают только на платформах, которые поддерживают следующие наборы требований.</span><span class="sxs-lookup"><span data-stu-id="047bb-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="047bb-117">Дополнительные информацию о наборах требований и работе с ними см. в подразделе "Указание приложений [Office и требований к API".](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="047bb-118">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="047bb-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="047bb-119">Поведение настраиваемой контекстной вкладки</span><span class="sxs-lookup"><span data-stu-id="047bb-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="047bb-120">Пользовательский интерфейс настраиваемой контекстной вкладки следует шаблону встроенных контекстных вкладок Office.</span><span class="sxs-lookup"><span data-stu-id="047bb-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="047bb-121">Ниже основных принципов размещения настраиваемой контекстной вкладки:</span><span class="sxs-lookup"><span data-stu-id="047bb-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="047bb-122">Когда пользовательская контекстная вкладка отображается, она отображается в правой части ленты.</span><span class="sxs-lookup"><span data-stu-id="047bb-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="047bb-123">Если одна или несколько встроенных контекстных вкладок и одна или несколько настраиваемые контекстные вкладки из надстроек видны одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="047bb-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="047bb-124">Если надстройка имеет несколько контекстных вкладок и существуют контексты, в которых отображается несколько из них, они отображаются в том порядке, в котором они определены в надстройке.</span><span class="sxs-lookup"><span data-stu-id="047bb-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="047bb-125">(Направление в том же направлении, что и язык Office, то есть направление слева направо на языках слева направо, а направление справа налево — на языках справа налево.) Подробные [сведения о том,](#define-the-groups-and-controls-that-appear-on-the-tab) как их определить, см. в поднаборе "Определение групп и элементов управления, которые отображаются на вкладке".</span><span class="sxs-lookup"><span data-stu-id="047bb-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="047bb-126">Если несколько надстроек имеет контекстную вкладку, которая отображается в определенном контексте, они отображаются в том порядке, в котором были запущены надстройки.</span><span class="sxs-lookup"><span data-stu-id="047bb-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="047bb-127">Настраиваемые *контекстные* вкладки, в отличие от настраиваемой основной вкладки, не добавляются окончательно на ленту приложения Office.</span><span class="sxs-lookup"><span data-stu-id="047bb-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="047bb-128">Они присутствуют только в документах Office, в которых работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="047bb-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="047bb-129">Основные действия по добавлению контекстной вкладки в надстройку</span><span class="sxs-lookup"><span data-stu-id="047bb-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="047bb-130">Далее приводится основной этап добавления настраиваемой контекстной вкладки в надстройку.</span><span class="sxs-lookup"><span data-stu-id="047bb-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="047bb-131">Настройте надстройку для использования общей времени работы.</span><span class="sxs-lookup"><span data-stu-id="047bb-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="047bb-132">Определите вкладку, группы и элементы управления, которые отображаются на ней.</span><span class="sxs-lookup"><span data-stu-id="047bb-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="047bb-133">Зарегистрируйте контекстную вкладку в Office.</span><span class="sxs-lookup"><span data-stu-id="047bb-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="047bb-134">Укажите условия, в которые вкладка будет видна.</span><span class="sxs-lookup"><span data-stu-id="047bb-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="047bb-135">Настройка надстройки для использования общей времени работы</span><span class="sxs-lookup"><span data-stu-id="047bb-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="047bb-136">Чтобы добавить настраиваемые контекстные вкладки, надстройка будет использовать общую времени работы.</span><span class="sxs-lookup"><span data-stu-id="047bb-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="047bb-137">Дополнительные сведения см. в настройках [надстройки для использования общей времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-137">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="047bb-138">Определение групп и элементов управления, которые отображаются на вкладке</span><span class="sxs-lookup"><span data-stu-id="047bb-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="047bb-139">В отличие от настраиваемой основной вкладки, которые определены с помощью XML в манифесте, настраиваемые контекстные вкладки определяются во время работы с BLOB JSON.</span><span class="sxs-lookup"><span data-stu-id="047bb-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="047bb-140">Ваш код разбрасирует большой объект в объект JavaScript, а затем передает объект [методу Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="047bb-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="047bb-141">Настраиваемые контекстные вкладки присутствуют только в документах, в которых в настоящее время работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="047bb-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="047bb-142">Это отличается от настраиваемой основной вкладки, которые добавляются на ленту приложения Office при установке надстройки и остаются на этом после открытия другого документа.</span><span class="sxs-lookup"><span data-stu-id="047bb-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="047bb-143">Кроме того, `requestCreateControls` метод можно запустить только один раз в сеансе надстройки.</span><span class="sxs-lookup"><span data-stu-id="047bb-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="047bb-144">Если он будет вызван повторно, будет выброшена ошибка.</span><span class="sxs-lookup"><span data-stu-id="047bb-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="047bb-145">Структура свойств и подэлементов BLOB-объекта JSON (и имен ключей) приблизительно параллельна структуре элемента [CustomTab](../reference/manifest/customtab.md) и его потомков в XML манифеста.</span><span class="sxs-lookup"><span data-stu-id="047bb-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="047bb-146">Пошаговое создание примера контекстных вкладок JSON.</span><span class="sxs-lookup"><span data-stu-id="047bb-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="047bb-147">(Полная схема контекстной вкладки JSON находится [вdynamic-ribbon.schema.js.](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)</span><span class="sxs-lookup"><span data-stu-id="047bb-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="047bb-148">Эта ссылка может не работать в период предварительного просмотра для контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="047bb-148">This link may not be working in the preview period for contextual tabs.</span></span> <span data-ttu-id="047bb-149">Если ссылка не работает, вы можете найти последний черновик схемы на черновике dynamic-ribbon.schema.js[на](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) Если вы работаете в Visual Studio Code, этот файл можно использовать для получения IntelliSense проверки JSON.</span><span class="sxs-lookup"><span data-stu-id="047bb-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="047bb-150">Дополнительные сведения см. в редактировании [JSON с помощью Visual Studio Code — схемы и параметры JSON.](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)</span><span class="sxs-lookup"><span data-stu-id="047bb-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="047bb-151">Начните с создания строки JSON с двумя свойствами массива с именем `actions` и `tabs` .</span><span class="sxs-lookup"><span data-stu-id="047bb-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="047bb-152">Массив — это спецификация всех функций, которые можно выполнять с помощью элементов управления `actions` на контекстной вкладке. Массив определяет одну или несколько контекстных вкладок `tabs` до *20*.</span><span class="sxs-lookup"><span data-stu-id="047bb-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="047bb-153">Этот простой пример контекстной вкладки будет иметь только одну кнопку и, следовательно, только одно действие.</span><span class="sxs-lookup"><span data-stu-id="047bb-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="047bb-154">Добавьте следующий в качестве единственный член `actions` массива.</span><span class="sxs-lookup"><span data-stu-id="047bb-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="047bb-155">Обратите внимание на эту разметку:</span><span class="sxs-lookup"><span data-stu-id="047bb-155">About this markup, note:</span></span>

    - <span data-ttu-id="047bb-156">Свойства `id` являются `type` обязательными.</span><span class="sxs-lookup"><span data-stu-id="047bb-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="047bb-157">Значением `type` может быть ExecuteFunction или ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="047bb-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="047bb-158">Свойство `functionName` используется только в том случае, если `type` значением является `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="047bb-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="047bb-159">Это имя функции, определенной в FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="047bb-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="047bb-160">Дополнительные сведения о FunctionFile см. в основных понятиях для команд [надстройки.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="047bb-161">На более позднем этапе вы соберем это действие с кнопкой на контекстной вкладке.</span><span class="sxs-lookup"><span data-stu-id="047bb-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="047bb-162">Добавьте следующий в качестве единственный член `tabs` массива.</span><span class="sxs-lookup"><span data-stu-id="047bb-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="047bb-163">Обратите внимание на эту разметку:</span><span class="sxs-lookup"><span data-stu-id="047bb-163">About this markup, note:</span></span>

    - <span data-ttu-id="047bb-164">Свойство `id` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="047bb-164">The `id` property is required.</span></span> <span data-ttu-id="047bb-165">Используйте краткий описательный ИД, уникальный для всех контекстных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="047bb-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="047bb-166">Свойство `label` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="047bb-166">The `label` property is required.</span></span> <span data-ttu-id="047bb-167">Это пользовательская строка, которая служит меткой контекстной вкладки.</span><span class="sxs-lookup"><span data-stu-id="047bb-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="047bb-168">Свойство `groups` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="047bb-168">The `groups` property is required.</span></span> <span data-ttu-id="047bb-169">Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен иметь по крайней мере один член *и не более 20*.</span><span class="sxs-lookup"><span data-stu-id="047bb-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="047bb-170">(Кроме того, существуют ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, а также количество групп.</span><span class="sxs-lookup"><span data-stu-id="047bb-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="047bb-171">Дополнительные сведения см. в следующем шаге.)</span><span class="sxs-lookup"><span data-stu-id="047bb-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="047bb-172">Кроме того, у объекта tab может быть необязательное свойство, которое указывает, отображается ли вкладка сразу после начала `visible` надстройки.</span><span class="sxs-lookup"><span data-stu-id="047bb-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="047bb-173">Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не активирует их видимость (например, пользователь выбирает сущность того или иного типа в документе), свойство по умолчанию имеет значение, когда его `visible` `false` нет.</span><span class="sxs-lookup"><span data-stu-id="047bb-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="047bb-174">В более позднем разделе мы покажем, как настроить свойство в `true` ответ на событие.</span><span class="sxs-lookup"><span data-stu-id="047bb-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="047bb-175">В простом примере контекстная вкладка имеет только одну группу.</span><span class="sxs-lookup"><span data-stu-id="047bb-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="047bb-176">Добавьте следующий в качестве единственный член `groups` массива.</span><span class="sxs-lookup"><span data-stu-id="047bb-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="047bb-177">Обратите внимание на эту разметку:</span><span class="sxs-lookup"><span data-stu-id="047bb-177">About this markup, note:</span></span>

    - <span data-ttu-id="047bb-178">Все свойства являются обязательной.</span><span class="sxs-lookup"><span data-stu-id="047bb-178">All the properties are required.</span></span>
    - <span data-ttu-id="047bb-179">Свойство должно быть уникальным для всех групп на `id` вкладке. Используйте краткий и описательный ИД.</span><span class="sxs-lookup"><span data-stu-id="047bb-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="047bb-180">Это `label` пользовательская строка, которая будет служить меткой группы.</span><span class="sxs-lookup"><span data-stu-id="047bb-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="047bb-181">Значение свойства — это массив объектов, которые указывают значки, которые будут иметься в группе на ленте в зависимости от размера ленты и окна `icon` приложения Office.</span><span class="sxs-lookup"><span data-stu-id="047bb-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="047bb-182">Значение свойства — это массив объектов, которые указывают кнопки и меню `controls` в группе.</span><span class="sxs-lookup"><span data-stu-id="047bb-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="047bb-183">Должно быть по крайней мере одно.</span><span class="sxs-lookup"><span data-stu-id="047bb-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="047bb-184">*Общее число элементов управления на всей вкладке не может быть больше 20.*</span><span class="sxs-lookup"><span data-stu-id="047bb-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="047bb-185">Например, можно иметь 3 группы с по 6 элементов управления и четвертую группу с 2 элементами управления, но нельзя иметь 4 группы по 6 элементов управления.</span><span class="sxs-lookup"><span data-stu-id="047bb-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. <span data-ttu-id="047bb-186">Каждая группа должна иметь значок размером не менее двух размеров: 32x32 пк и 80x80 пк.</span><span class="sxs-lookup"><span data-stu-id="047bb-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="047bb-187">Кроме того, можно использовать значки размером 16x16 пк, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="047bb-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="047bb-188">Office определяет, какой значок использовать в зависимости от размера ленты и окна приложения Office.</span><span class="sxs-lookup"><span data-stu-id="047bb-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="047bb-189">Добавьте следующие объекты в массив значков.</span><span class="sxs-lookup"><span data-stu-id="047bb-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="047bb-190">(Если размер окна и ленты достаточно велик для  появления хотя бы одного из элементов управления в группе, значок группы вообще не отображается.</span><span class="sxs-lookup"><span data-stu-id="047bb-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="047bb-191">Например, просмотрите группу **стилей** на ленте Word при сжатии и расширении окна Word.) Обратите внимание на эту разметку:</span><span class="sxs-lookup"><span data-stu-id="047bb-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="047bb-192">Оба свойства являются обязательной.</span><span class="sxs-lookup"><span data-stu-id="047bb-192">Both the properties are required.</span></span>
    - <span data-ttu-id="047bb-193">Единица `size` измерения свойства — пиксели.</span><span class="sxs-lookup"><span data-stu-id="047bb-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="047bb-194">Значки всегда квадратные, поэтому числом является и высота, и ширина.</span><span class="sxs-lookup"><span data-stu-id="047bb-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="047bb-195">Свойство `sourceLocation` указывает полный URL-адрес значка.</span><span class="sxs-lookup"><span data-stu-id="047bb-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="047bb-196">Так же, как обычно необходимо изменить URL-адреса в манифесте надстройки при переходе от разработки к производственной (например, при изменении домена с localhost на contoso.com), необходимо также изменить URL-адреса в контекстных вкладок JSON.</span><span class="sxs-lookup"><span data-stu-id="047bb-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. <span data-ttu-id="047bb-197">В нашем простом примере группа имеет только одну кнопку.</span><span class="sxs-lookup"><span data-stu-id="047bb-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="047bb-198">Добавьте следующий объект в качестве единственный член `controls` массива.</span><span class="sxs-lookup"><span data-stu-id="047bb-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="047bb-199">Обратите внимание на эту разметку:</span><span class="sxs-lookup"><span data-stu-id="047bb-199">About this markup, note:</span></span>

    - <span data-ttu-id="047bb-200">Все свойства, кроме `enabled` , являются обязательной.</span><span class="sxs-lookup"><span data-stu-id="047bb-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="047bb-201">`type` указывает тип управления.</span><span class="sxs-lookup"><span data-stu-id="047bb-201">`type` specifies the type of control.</span></span> <span data-ttu-id="047bb-202">Значениями могут быть "Button", "Menu" или "MobileButton".</span><span class="sxs-lookup"><span data-stu-id="047bb-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="047bb-203">`id` может быть до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="047bb-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="047bb-204">`actionId` должен быть ИД действия, определенного в `actions` массиве.</span><span class="sxs-lookup"><span data-stu-id="047bb-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="047bb-205">(См. шаг 1 этого раздела.)</span><span class="sxs-lookup"><span data-stu-id="047bb-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="047bb-206">`label` — это пользовательская строка, которая служит меткой кнопки.</span><span class="sxs-lookup"><span data-stu-id="047bb-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="047bb-207">`superTip` представляет собой форматную форму подсказки.</span><span class="sxs-lookup"><span data-stu-id="047bb-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="047bb-208">Необходимы `title` и `description` свойства, и свойства.</span><span class="sxs-lookup"><span data-stu-id="047bb-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="047bb-209">`icon` указывает значки для кнопки.</span><span class="sxs-lookup"><span data-stu-id="047bb-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="047bb-210">Здесь также применимы предыдущие замечания о значке группы.</span><span class="sxs-lookup"><span data-stu-id="047bb-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="047bb-211">`enabled` (необязательно) указывает, включена ли кнопка при отжатии контекстной вкладки.</span><span class="sxs-lookup"><span data-stu-id="047bb-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="047bb-212">Значение по умолчанию , если нет `true` .</span><span class="sxs-lookup"><span data-stu-id="047bb-212">The default if not present is `true`.</span></span> 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
<span data-ttu-id="047bb-213">Ниже приводится полный пример BLOB-blob JSON:</span><span class="sxs-lookup"><span data-stu-id="047bb-213">The following is the complete example of the JSON blob:</span></span>

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="047bb-214">Регистрация контекстной вкладки в Office с помощью requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="047bb-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="047bb-215">Контекстная вкладка регистрируется в Office путем вызова метода [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="047bb-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="047bb-216">Обычно это делается в функции, назначенной методу, `Office.initialize` или с помощью этого `Office.onReady` метода.</span><span class="sxs-lookup"><span data-stu-id="047bb-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="047bb-217">Подробнее об этих методах и инициализации надстройки см. в инициализации [надстройки Office.](../develop/initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="047bb-218">Однако вы можете вызвать метод в любое время после инициализации.</span><span class="sxs-lookup"><span data-stu-id="047bb-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="047bb-219">Метод `requestCreateControls` может быть вызван только один раз в заданном сеансе надстройки.</span><span class="sxs-lookup"><span data-stu-id="047bb-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="047bb-220">Если она будет вызвана повторно, будет выброшена ошибка.</span><span class="sxs-lookup"><span data-stu-id="047bb-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="047bb-221">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="047bb-221">The following is an example.</span></span> <span data-ttu-id="047bb-222">Обратите внимание, что перед тем как передать строку JSON в функцию JavaScript, ее необходимо преобразовать в объект JavaScript с помощью `JSON.parse` метода.</span><span class="sxs-lookup"><span data-stu-id="047bb-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="047bb-223">Укажите контексты, когда вкладка будет видна с помощью requestUpdate</span><span class="sxs-lookup"><span data-stu-id="047bb-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="047bb-224">Как правило, настраиваемая контекстная вкладка должна отображаться, когда событие, инициированное пользователем, изменяет контекст надстройки.</span><span class="sxs-lookup"><span data-stu-id="047bb-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="047bb-225">Рассмотрим сценарий, в котором вкладка должна быть видна, когда активируется диаграмма (на стандартной таблице книги Excel).</span><span class="sxs-lookup"><span data-stu-id="047bb-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="047bb-226">Начните с назначения обработчиков.</span><span class="sxs-lookup"><span data-stu-id="047bb-226">Begin by assigning handlers.</span></span> <span data-ttu-id="047bb-227">Обычно это делается в методе, как в следующем примере, который назначает обработчики (созданные на более позднем этапе) событиям всех диаграмм на `Office.onReady` `onActivated` этом `onDeactivated` графике.</span><span class="sxs-lookup"><span data-stu-id="047bb-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

<span data-ttu-id="047bb-228">Затем определите обработчики.</span><span class="sxs-lookup"><span data-stu-id="047bb-228">Next, define the handlers.</span></span> <span data-ttu-id="047bb-229">Ниже приводится простой пример ошибки `showDataTab` [HostRestartNeeded,](#handle-the-hostrestartneeded-error) но более надежную версию функции см. далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="047bb-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="047bb-230">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="047bb-230">About this code, note:</span></span>

- <span data-ttu-id="047bb-231">Office определяет время обновления состояния ленты.</span><span class="sxs-lookup"><span data-stu-id="047bb-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="047bb-232">Метод  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) очереди запрос на обновление.</span><span class="sxs-lookup"><span data-stu-id="047bb-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="047bb-233">Метод разрешит объект сразу после того, как он задюет запрос в очередь, а не при `Promise` обновлении ленты.</span><span class="sxs-lookup"><span data-stu-id="047bb-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="047bb-234">Параметром метода является объект `requestUpdate` [RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) указывает вкладку по ее ИД точно так же, как указано в *JSON* и (2) определяет видимость вкладки.</span><span class="sxs-lookup"><span data-stu-id="047bb-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="047bb-235">Если имеется несколько настраиваемой контекстной вкладки, которая должна быть видна в одном контексте, в массив просто добавляются дополнительные объекты `tabs` вкладок.</span><span class="sxs-lookup"><span data-stu-id="047bb-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

<span data-ttu-id="047bb-236">Обработок для скрытие вкладки почти идентичен, за исключением того, что он возвращает `visible` `false` свойство.</span><span class="sxs-lookup"><span data-stu-id="047bb-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="047bb-237">Библиотека JavaScript для Office также предоставляет несколько интерфейсов (типов), упрощая создание `RibbonUpdateData` объекта.</span><span class="sxs-lookup"><span data-stu-id="047bb-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="047bb-238">Ниже приводится функция `showDataTab` в TypeScript, которая использует эти типы.</span><span class="sxs-lookup"><span data-stu-id="047bb-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="047bb-239">Одновременное переключеть видимость вкладки и состояние включенной кнопки</span><span class="sxs-lookup"><span data-stu-id="047bb-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="047bb-240">Этот метод также используется для включения или отключения состояния настраиваемой кнопки на настраиваемой контекстной вкладке или в пользовательской `requestUpdate` основной вкладке. Дополнительные сведения см. в подстройке "Включить и отключить [команды надстройки".](disable-add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="047bb-241">В ряде сценариев может потребоваться одновременное изменение видимости вкладки и состояния включенной кнопки.</span><span class="sxs-lookup"><span data-stu-id="047bb-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="047bb-242">Это можно сделать одним вызовом `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="047bb-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="047bb-243">Ниже приводится пример, в котором кнопка на основной вкладке включена одновременно с видимой контекстной вкладками.</span><span class="sxs-lookup"><span data-stu-id="047bb-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
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
            ]}
        ]
    });
}
```

<span data-ttu-id="047bb-244">В следующем примере включенная кнопка находится на той же контекстной вкладке, которая отображается.</span><span class="sxs-lookup"><span data-stu-id="047bb-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="047bb-245">Локализация BLOB JSON</span><span class="sxs-lookup"><span data-stu-id="047bb-245">Localizing the JSON blob</span></span>

<span data-ttu-id="047bb-246">Передаваемый BLOB-проект JSON не локализуется так же, как локализована разметка манифеста для настраиваемой основной вкладки (что описано в локализации control из `requestCreateControls` манифеста). [](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="047bb-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="047bb-247">Вместо этого локализация должна происходить во время работы с использованием отдельных BLOB-ок JSON для каждого из региональных стандартов.</span><span class="sxs-lookup"><span data-stu-id="047bb-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="047bb-248">Мы рекомендуем использовать заявление, которое тестирует `switch` [свойство Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage)</span><span class="sxs-lookup"><span data-stu-id="047bb-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="047bb-249">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="047bb-249">The following is an example:</span></span>

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

<span data-ttu-id="047bb-250">Затем код вызывает функцию, чтобы получить локализованный BLOB-код, который передается , как в `requestCreateControls` следующем примере:</span><span class="sxs-lookup"><span data-stu-id="047bb-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="047bb-251">Best practices for custom contextual tabs</span><span class="sxs-lookup"><span data-stu-id="047bb-251">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="047bb-252">Реализация альтернативного интерфейса пользователя, когда настраиваемые контекстные вкладки не поддерживаются</span><span class="sxs-lookup"><span data-stu-id="047bb-252">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="047bb-253">Некоторые сочетания платформы, приложения Office и сборки Office не `requestCreateControls` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="047bb-253">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="047bb-254">Ваша надстройка должна быть разработана для предоставления пользователям, которые запускают надстройки в одной из этих комбинаций.</span><span class="sxs-lookup"><span data-stu-id="047bb-254">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="047bb-255">В следующих разделах описаны два способа обеспечения отката.</span><span class="sxs-lookup"><span data-stu-id="047bb-255">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="047bb-256">Использование неконтекстуальных вкладок или элементов управления</span><span class="sxs-lookup"><span data-stu-id="047bb-256">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="047bb-257">Существует элемент манифеста [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)предназначенный для создания отката в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка запущена в приложении или платформе, не поддерживаю которой настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="047bb-257">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="047bb-258">Простейшая стратегия использования этого элемента заключается в *том,* что вы определяете в манифесте одну или несколько настраиваемых ядер вкладок (то есть неконтекстуальных настраиваемых вкладок), дублирующих настройки ленты настраиваемых контекстных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="047bb-258">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="047bb-259">Но вы `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` добавляете в качестве первого потомка элемента [CustomTab.](../reference/manifest/customtab.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-259">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="047bb-260">Это может быть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="047bb-260">The effect of doing so is the following:</span></span>

- <span data-ttu-id="047bb-261">Если надстройка работает на приложениях и платформах, поддерживаюх настраиваемые контекстные вкладки, настраиваемая основная вкладка не будет отображаться на ленте.</span><span class="sxs-lookup"><span data-stu-id="047bb-261">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="047bb-262">Вместо этого настраиваемая контекстная вкладка будет создана, когда надстройка вызывает `requestCreateControls` метод.</span><span class="sxs-lookup"><span data-stu-id="047bb-262">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="047bb-263">Если надстройка работает в приложении или платформе, которые не поддерживаются, на ленте появляется настраиваемая вкладка  `requestCreateControls` "Ядро".</span><span class="sxs-lookup"><span data-stu-id="047bb-263">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="047bb-264">Ниже приводится пример этой простой стратегии.</span><span class="sxs-lookup"><span data-stu-id="047bb-264">The following is an example of this simple strategy.</span></span>

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
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="047bb-265">Эта простая стратегия использует настраиваемую основную вкладку, которая зеркально отражает настраиваемую контекстную вкладку с ее потомками и группами элементов управления, но вы можете использовать более сложную стратегию.</span><span class="sxs-lookup"><span data-stu-id="047bb-265">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="047bb-266">Этот элемент также можно добавить как (первый) в элементы Group и Control (тип кнопки и `<OverriddenByRibbonApi>` [тип](../reference/manifest/control.md#menu-dropdown-button-controls)меню) и элементы [](../reference/manifest/group.md) [](../reference/manifest/control.md) [](../reference/manifest/control.md#button-control) `<Item>` меню.</span><span class="sxs-lookup"><span data-stu-id="047bb-266">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="047bb-267">Это позволяет распределять группы и элементы управления, которые в противном случае отображаются на контекстной вкладке между различными группами, кнопками и меню на различных настраиваемой вкладке ядра.</span><span class="sxs-lookup"><span data-stu-id="047bb-267">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="047bb-268">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="047bb-268">The following is an example.</span></span> <span data-ttu-id="047bb-269">Обратите внимание, что MyButton будет отображаться на настраиваемой основной вкладке только в том случае, если настраиваемые контекстные вкладки не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="047bb-269">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="047bb-270">Однако родительская группа и настраиваемая основная вкладка будут отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="047bb-270">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="047bb-271">Дополнительные примеры см. в [подразделе OverriddenByRibbonApi.](../reference/manifest/overriddenbyribbonapi.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-271">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="047bb-272">Если родительская вкладка, группа или меню помечены , то она не отображается, а вся ее родительская разметка игнорируется, если настраиваемые контекстные вкладки не `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="047bb-272">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="047bb-273">Поэтому не имеет значения, имеет ли какой-либо из этих элементов элемент `<OverriddenByRibbonApi>` или его значение.</span><span class="sxs-lookup"><span data-stu-id="047bb-273">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="047bb-274">Это означает, что если элемент меню, элемент управления или группа должны быть видимыми во всех контекстах, то не только не должен быть помечен, но и его предок меню, группа и вкладка также не должны помечаться таким `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` образом. </span><span class="sxs-lookup"><span data-stu-id="047bb-274">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="047bb-275">Не *пометить* все элементы вкладки, группы или меню с помощью `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="047bb-275">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="047bb-276">Это не имеет смысла, если родительский элемент помечен по причинам, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` заданным в предыдущем абзаце.</span><span class="sxs-lookup"><span data-stu-id="047bb-276">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="047bb-277">Кроме того, если не использовать родительский текст (или установить для него его), то родительский будет отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки, но он будет пустым при их `<OverriddenByRibbonApi>` `false` поддержке.</span><span class="sxs-lookup"><span data-stu-id="047bb-277">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="047bb-278">Таким образом, если при поддержке настраиваемой контекстной вкладки не должны отображаться все эти элементы, пометите родительский элемент и только родительский элемент с помощью `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="047bb-278">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="047bb-279">Использование API,которые показывают или скрывают области задач в указанных контекстах</span><span class="sxs-lookup"><span data-stu-id="047bb-279">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="047bb-280">В качестве альтернативы надстройка может определить области задач с помощью элементов управления пользовательского интерфейса, дублирующих функции элементов управления на `<OverriddenByRibbonApi>` настраиваемой контекстной вкладке. Затем используйте методы [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) и [Office.addin.hide,](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) чтобы показать область задач, когда и только когда контекстная вкладка была бы показана, если бы она была поддерживается.</span><span class="sxs-lookup"><span data-stu-id="047bb-280">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="047bb-281">Дополнительные сведения об использовании этих методов см. в подстройке "Показать" или "Скрыть" области [задач надстройки Office.](../develop/show-hide-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="047bb-281">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="047bb-282">Обработка ошибки HostRestartNeeded</span><span class="sxs-lookup"><span data-stu-id="047bb-282">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="047bb-283">В некоторых случаях Office не может обновить ленту и возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="047bb-283">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="047bb-284">Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть.</span><span class="sxs-lookup"><span data-stu-id="047bb-284">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="047bb-285">Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="047bb-285">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="047bb-286">Код должен обработать эту ошибку.</span><span class="sxs-lookup"><span data-stu-id="047bb-286">Your code should handle this error.</span></span> <span data-ttu-id="047bb-287">Ниже приводится пример того, как это сделать.</span><span class="sxs-lookup"><span data-stu-id="047bb-287">The following is an example of how.</span></span> <span data-ttu-id="047bb-288">В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.</span><span class="sxs-lookup"><span data-stu-id="047bb-288">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
