---
title: Создание настраиваемой контекстной вкладки Office надстроек
description: Узнайте, как добавить настраиваемые контекстные вкладки в Office надстройку.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 980beb24a3d384ecf21da44db288272a1ab1b0e3
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330173"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="ba4df-103">Создание настраиваемой контекстной вкладки Office надстроек</span><span class="sxs-lookup"><span data-stu-id="ba4df-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="ba4df-104">Контекстная вкладка — это скрытый контроль вкладок в ленте Office, отображаемой в строке вкладок, когда указанное событие происходит в Office документе.</span><span class="sxs-lookup"><span data-stu-id="ba4df-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="ba4df-105">Например, **вкладка "Дизайн** таблицы", которая отображается на Excel при выборе таблицы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="ba4df-106">Вы можете включить настраиваемые контекстные вкладки в Office надстройки и указать, когда они видны или скрыты, создав обработчики событий, которые изменяют видимость.</span><span class="sxs-lookup"><span data-stu-id="ba4df-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="ba4df-107">(Однако настраиваемые контекстные вкладки не реагируют на изменения фокуса.)</span><span class="sxs-lookup"><span data-stu-id="ba4df-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="ba4df-108">В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией.</span><span class="sxs-lookup"><span data-stu-id="ba4df-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="ba4df-109">Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.</span><span class="sxs-lookup"><span data-stu-id="ba4df-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="ba4df-110">Основные концепции команд надстроек</span><span class="sxs-lookup"><span data-stu-id="ba4df-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="ba4df-111">Пользовательские контекстные вкладки в настоящее время поддерживаются только на Excel и только на этих платформах и сборках:</span><span class="sxs-lookup"><span data-stu-id="ba4df-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="ba4df-112">Excel на Windows (только Microsoft 365 подписка): Версия 2102 (сборка 13801.20294) или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="ba4df-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="ba4df-113">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="ba4df-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="ba4df-114">Настраиваемые контекстные вкладки работают только на платформах, поддерживаюх следующие наборы требований.</span><span class="sxs-lookup"><span data-stu-id="ba4df-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="ba4df-115">Дополнительные подробности о наборах требований и работе с ними см. в Office [приложений и API.](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="ba4df-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="ba4df-116">RibbonApi 1.2</span><span class="sxs-lookup"><span data-stu-id="ba4df-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="ba4df-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="ba4df-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="ba4df-118">Вы можете использовать проверки времени запуска в коде, чтобы проверить, поддерживает ли комбинация хост и платформа пользователя эти наборы требований, описанные в описании Office приложений и [требований API.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="ba4df-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="ba4df-119">(Метод указания наборов требований в манифесте, который также описан в этой статье, в настоящее время не работает для RibbonApi 1.2.) Кроме того, вы можете [реализовать альтернативный интерфейс интерфейса, если пользовательские](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)контекстные вкладки не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="ba4df-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="ba4df-120">Поведение пользовательских контекстных вкладок</span><span class="sxs-lookup"><span data-stu-id="ba4df-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="ba4df-121">Пользовательский интерфейс для пользовательских контекстных вкладок следует шаблону встроенных Office контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="ba4df-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="ba4df-122">Основные принципы размещения пользовательских контекстных вкладок:</span><span class="sxs-lookup"><span data-stu-id="ba4df-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="ba4df-123">Когда отображается настраиваемая контекстная вкладка, она отображается на правом конце ленты.</span><span class="sxs-lookup"><span data-stu-id="ba4df-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="ba4df-124">Если одна или несколько встроенных контекстных вкладок и одна или несколько пользовательских контекстных вкладок из надстроек видны одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="ba4df-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="ba4df-125">Если надстройка имеет несколько контекстных вкладок и есть контексты, в которых видно несколько, они отображаются в порядке, в котором они определены в вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="ba4df-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="ba4df-126">(Это направление в том же направлении, что и язык Office, то есть слева направо на левом и правом языках, но справа налево на языках справа налево.) Сведения [о том,](#define-the-groups-and-controls-that-appear-on-the-tab) как их определить, см. в материале Определение групп и элементов управления, которые отображаются на вкладке.</span><span class="sxs-lookup"><span data-stu-id="ba4df-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="ba4df-127">Если несколько надстроек имеет контекстную вкладку, которая видна в определенном контексте, они отображаются в порядке запуска надстроек.</span><span class="sxs-lookup"><span data-stu-id="ba4df-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="ba4df-128">Настраиваемые *контекстные* вкладки, в отличие от настраиваемой основной вкладки, не добавляются Office ленту приложения.</span><span class="sxs-lookup"><span data-stu-id="ba4df-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="ba4df-129">Они присутствуют только в Office документах, на которых работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="ba4df-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="ba4df-130">Основные действия по включаемой контекстной вкладке в надстройку</span><span class="sxs-lookup"><span data-stu-id="ba4df-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="ba4df-131">Следующие основные действия для добавления настраиваемой контекстной вкладки в надстройку:</span><span class="sxs-lookup"><span data-stu-id="ba4df-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="ba4df-132">Настройте надстройку для использования общего времени запуска.</span><span class="sxs-lookup"><span data-stu-id="ba4df-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="ba4df-133">Определите вкладку, группы и элементы управления, которые отображаются на ней.</span><span class="sxs-lookup"><span data-stu-id="ba4df-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="ba4df-134">Зарегистрируйте контекстную вкладку с помощью Office.</span><span class="sxs-lookup"><span data-stu-id="ba4df-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="ba4df-135">Укажите обстоятельства, когда вкладка будет видна.</span><span class="sxs-lookup"><span data-stu-id="ba4df-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="ba4df-136">Настройка надстройки для использования общего времени работы</span><span class="sxs-lookup"><span data-stu-id="ba4df-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="ba4df-137">Добавление настраиваемой контекстной вкладки требует от надстройки использовать общее время работы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="ba4df-138">Дополнительные сведения см. в [раздел Настройка надстройки для использования общего времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="ba4df-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="ba4df-139">Определение групп и элементов управления, которые отображаются на вкладке</span><span class="sxs-lookup"><span data-stu-id="ba4df-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="ba4df-140">В отличие от настраиваемой вкладки ядра, которые определяются с помощью XML в манифесте, настраиваемые контекстные вкладки определяются во время запуска с помощью BLOB JSON.</span><span class="sxs-lookup"><span data-stu-id="ba4df-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="ba4df-141">Код разрезает blob в объект JavaScript, а затем передает объект [методу Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="ba4df-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="ba4df-142">Настраиваемые контекстные вкладки присутствуют только в документах, на которых в настоящее время запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="ba4df-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="ba4df-143">Это отличается от настраиваемой основной вкладки, которые добавляются в ленту Office приложения при установке надстройки и остаются в момент открытия другого документа.</span><span class="sxs-lookup"><span data-stu-id="ba4df-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="ba4df-144">Кроме того, `requestCreateControls` метод можно запускать только один раз в сеансе надстройки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="ba4df-145">Если он снова вызван, ошибка будет выброшена.</span><span class="sxs-lookup"><span data-stu-id="ba4df-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="ba4df-146">Структура свойств и свойств BLOB JSON (и имен ключей) примерно параллельна структуре элемента [CustomTab](../reference/manifest/customtab.md) и его элементов потомка в манифесте XML.</span><span class="sxs-lookup"><span data-stu-id="ba4df-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="ba4df-147">Мы пошаговую соберем пример контекстных вкладок JSON blob.</span><span class="sxs-lookup"><span data-stu-id="ba4df-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="ba4df-148">(Полная схема контекстной вкладки JSON находится [вdynamic-ribbon.schema.js.](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)</span><span class="sxs-lookup"><span data-stu-id="ba4df-148">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="ba4df-149">Эта ссылка может не работать в период предварительного просмотра для контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="ba4df-149">This link may not be working in the preview period for contextual tabs.</span></span> <span data-ttu-id="ba4df-150">Если ссылка не работает, вы можете найти последний черновик схемы на черновике dynamic-ribbon.schema.js[на](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) Если вы работаете в Visual Studio Code, вы можете использовать этот файл для получения IntelliSense и проверки JSON.</span><span class="sxs-lookup"><span data-stu-id="ba4df-150">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="ba4df-151">Дополнительные сведения см. в [статью Редактирование JSON с Visual Studio Code - схемы и параметры JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="ba4df-151">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="ba4df-152">Начните с создания строки JSON с двумя свойствами массива с `actions` именем и `tabs` .</span><span class="sxs-lookup"><span data-stu-id="ba4df-152">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="ba4df-153">Массив — это спецификация всех функций, которые можно выполнять с помощью `actions` элементов управления на контекстной вкладке. Массив определяет одну или несколько контекстных вкладок, не более `tabs` *20*.</span><span class="sxs-lookup"><span data-stu-id="ba4df-153">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="ba4df-154">Этот простой пример контекстной вкладки будет иметь только одну кнопку и, следовательно, только одно действие.</span><span class="sxs-lookup"><span data-stu-id="ba4df-154">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="ba4df-155">Добавьте следующее как единственный член `actions` массива.</span><span class="sxs-lookup"><span data-stu-id="ba4df-155">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="ba4df-156">Об этой разметки обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ba4df-156">About this markup, note:</span></span>

    - <span data-ttu-id="ba4df-157">Свойства `id` `type` и свойства обязательны.</span><span class="sxs-lookup"><span data-stu-id="ba4df-157">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="ba4df-158">Значение может `type` быть "ExecuteFunction" или "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="ba4df-158">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="ba4df-159">Свойство `functionName` используется только при значении `type` `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="ba4df-159">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="ba4df-160">Это имя функции, определенной в FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="ba4df-160">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="ba4df-161">Дополнительные сведения о FunctionFile см. в базовых [понятиях команд надстройки.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="ba4df-161">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="ba4df-162">На более позднем этапе вы соберете это действие на кнопку на вкладке contextual.</span><span class="sxs-lookup"><span data-stu-id="ba4df-162">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="ba4df-163">Добавьте следующее как единственный член `tabs` массива.</span><span class="sxs-lookup"><span data-stu-id="ba4df-163">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="ba4df-164">Об этой разметки обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ba4df-164">About this markup, note:</span></span>

    - <span data-ttu-id="ba4df-165">Свойство `id` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="ba4df-165">The `id` property is required.</span></span> <span data-ttu-id="ba4df-166">Используйте краткий описательный ID, уникальный среди всех контекстных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="ba4df-166">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="ba4df-167">Свойство `label` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="ba4df-167">The `label` property is required.</span></span> <span data-ttu-id="ba4df-168">Это удобное строка, которая служит меткой контекстной вкладки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-168">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="ba4df-169">Свойство `groups` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="ba4df-169">The `groups` property is required.</span></span> <span data-ttu-id="ba4df-170">Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен иметь по крайней мере один член *и не более 20*.</span><span class="sxs-lookup"><span data-stu-id="ba4df-170">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="ba4df-171">(Существует также ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, что также ограничивает количество групп, которые у вас есть.</span><span class="sxs-lookup"><span data-stu-id="ba4df-171">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="ba4df-172">Дополнительные сведения см. в следующем шаге.)</span><span class="sxs-lookup"><span data-stu-id="ba4df-172">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="ba4df-173">Объект вкладки также может иметь необязательное свойство, которое указывает, видна ли вкладка сразу после `visible` начала надстройки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-173">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="ba4df-174">Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не вызовет их видимость (например, если пользователь выбирает объект определенного типа в документе), свойство по умолчанию не будет `visible` `false` присутствовать.</span><span class="sxs-lookup"><span data-stu-id="ba4df-174">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="ba4df-175">В более позднем разделе мы покажем, как настроить свойство в ответ `true` на событие.</span><span class="sxs-lookup"><span data-stu-id="ba4df-175">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="ba4df-176">В простом непрерывном примере контекстная вкладка имеет только одну группу.</span><span class="sxs-lookup"><span data-stu-id="ba4df-176">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="ba4df-177">Добавьте следующее как единственный член `groups` массива.</span><span class="sxs-lookup"><span data-stu-id="ba4df-177">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="ba4df-178">Об этой разметки обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ba4df-178">About this markup, note:</span></span>

    - <span data-ttu-id="ba4df-179">Все свойства необходимы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-179">All the properties are required.</span></span>
    - <span data-ttu-id="ba4df-180">Свойство должно быть уникальным среди всех групп на `id` вкладке. Используйте краткий описательный ID.</span><span class="sxs-lookup"><span data-stu-id="ba4df-180">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="ba4df-181">Строка является удобной для `label` пользователя, которая служит в качестве метки группы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-181">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="ba4df-182">Значение свойства — массив объектов, которые указывают значки, которые будут иметься у группы на ленте в зависимости от размера ленты и `icon` окна Office приложения.</span><span class="sxs-lookup"><span data-stu-id="ba4df-182">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="ba4df-183">Значение свойства — это массив объектов, которые указывают кнопки и `controls` меню в группе.</span><span class="sxs-lookup"><span data-stu-id="ba4df-183">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="ba4df-184">Должно быть по крайней мере одно.</span><span class="sxs-lookup"><span data-stu-id="ba4df-184">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="ba4df-185">*Общее число элементов управления на всей вкладке может быть не более 20.*</span><span class="sxs-lookup"><span data-stu-id="ba4df-185">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="ba4df-186">Например, можно иметь 3 группы с 6 элементами управления и четвертую группу с 2 элементами управления, но нельзя иметь 4 группы с 6 элементами управления каждой.</span><span class="sxs-lookup"><span data-stu-id="ba4df-186">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="ba4df-187">Каждая группа должна иметь значок не менее двух размеров: 32x32 px и 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="ba4df-187">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="ba4df-188">Кроме того, можно использовать значки размеров 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="ba4df-188">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="ba4df-189">Office определяет, какой значок использовать в зависимости от размера ленты и Office окна приложения.</span><span class="sxs-lookup"><span data-stu-id="ba4df-189">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="ba4df-190">Добавьте следующие объекты в массив значок.</span><span class="sxs-lookup"><span data-stu-id="ba4df-190">Add the following objects to the icon array.</span></span> <span data-ttu-id="ba4df-191">(Если размеры окна и ленты достаточно большие для  появления хотя бы одного из элементов управления в группе, то не отображается значок группы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-191">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="ba4df-192">Например, просмотрите группу **Стилей** на ленте Word при сжатии и расширении окна Word.) Об этой разметки обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ba4df-192">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="ba4df-193">Необходимы оба свойства.</span><span class="sxs-lookup"><span data-stu-id="ba4df-193">Both the properties are required.</span></span>
    - <span data-ttu-id="ba4df-194">Единица `size` свойства измерения — пиксели.</span><span class="sxs-lookup"><span data-stu-id="ba4df-194">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="ba4df-195">Значки всегда квадратные, поэтому число — это как высота, так и ширина.</span><span class="sxs-lookup"><span data-stu-id="ba4df-195">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="ba4df-196">Свойство `sourceLocation` указывает полный URL-адрес значка.</span><span class="sxs-lookup"><span data-stu-id="ba4df-196">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="ba4df-197">Как правило, при переходе от разработки к производству (например, при изменении домена с локального на contoso.com) необходимо изменить URL-адреса в контекстных вкладок JSON.</span><span class="sxs-lookup"><span data-stu-id="ba4df-197">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="ba4df-198">В нашем простом непрерывном примере у группы есть только одна кнопка.</span><span class="sxs-lookup"><span data-stu-id="ba4df-198">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="ba4df-199">Добавьте следующий объект как единственный член `controls` массива.</span><span class="sxs-lookup"><span data-stu-id="ba4df-199">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="ba4df-200">Об этой разметки обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ba4df-200">About this markup, note:</span></span>

    - <span data-ttu-id="ba4df-201">Все свойства, за `enabled` исключением, необходимы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-201">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="ba4df-202">`type` указывает тип управления.</span><span class="sxs-lookup"><span data-stu-id="ba4df-202">`type` specifies the type of control.</span></span> <span data-ttu-id="ba4df-203">Значения могут быть "Button", "Menu" или "MobileButton".</span><span class="sxs-lookup"><span data-stu-id="ba4df-203">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="ba4df-204">`id` может быть до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="ba4df-204">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="ba4df-205">`actionId` должен быть ID действия, определенного в `actions` массиве.</span><span class="sxs-lookup"><span data-stu-id="ba4df-205">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="ba4df-206">(См. шаг 1 этого раздела.)</span><span class="sxs-lookup"><span data-stu-id="ba4df-206">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="ba4df-207">`label` является удобной строкой, которая служит в качестве метки кнопки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-207">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="ba4df-208">`superTip` представляет собой богатую форму подсказки инструмента.</span><span class="sxs-lookup"><span data-stu-id="ba4df-208">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="ba4df-209">Требуются `title` `description` как свойства, так и свойства.</span><span class="sxs-lookup"><span data-stu-id="ba4df-209">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="ba4df-210">`icon` указывает значки для кнопки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-210">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="ba4df-211">Предыдущие замечания о значке группы применяются и здесь.</span><span class="sxs-lookup"><span data-stu-id="ba4df-211">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="ba4df-212">`enabled` (необязательный) указывает, включена ли кнопка при запусках контекстной вкладки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-212">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="ba4df-213">Если по умолчанию `true` нет.</span><span class="sxs-lookup"><span data-stu-id="ba4df-213">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="ba4df-214">Ниже приводится полный пример BLOB JSON:</span><span class="sxs-lookup"><span data-stu-id="ba4df-214">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="ba4df-215">Регистрация контекстной вкладки с помощью Office с помощью requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="ba4df-215">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="ba4df-216">Контекстная вкладка регистрируется с помощью Office путем вызова [метода Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="ba4df-216">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="ba4df-217">Обычно это делается в функции, назначенной или `Office.initialize` с помощью `Office.onReady` метода.</span><span class="sxs-lookup"><span data-stu-id="ba4df-217">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="ba4df-218">Дополнительные данные об этих методах и инициализации надстройки см. в Office [надстройки.](../develop/initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="ba4df-218">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="ba4df-219">Однако вы можете вызвать метод в любое время после инициализации.</span><span class="sxs-lookup"><span data-stu-id="ba4df-219">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ba4df-220">Метод может быть вызван только один раз в `requestCreateControls` заданном сеансе надстройки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-220">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="ba4df-221">Ошибка будет выброшена, если она будет вызвана снова.</span><span class="sxs-lookup"><span data-stu-id="ba4df-221">An error is thrown if it is called again.</span></span>

<span data-ttu-id="ba4df-222">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="ba4df-222">The following is an example.</span></span> <span data-ttu-id="ba4df-223">Обратите внимание, что строка JSON должна быть преобразована в объект JavaScript с помощью метода, прежде чем она может быть передана `JSON.parse` функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ba4df-223">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="ba4df-224">Укажите контексты, когда вкладка будет видна с помощью requestUpdate</span><span class="sxs-lookup"><span data-stu-id="ba4df-224">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="ba4df-225">Как правило, настраиваемая контекстная вкладка должна отображаться, когда инициированное пользователем событие меняет контекст надстройки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-225">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="ba4df-226">Рассмотрим сценарий, в котором вкладка должна быть видна при активации диаграммы (по умолчанию в Excel книге).</span><span class="sxs-lookup"><span data-stu-id="ba4df-226">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="ba4df-227">Начните с назначения обработчиков.</span><span class="sxs-lookup"><span data-stu-id="ba4df-227">Begin by assigning handlers.</span></span> <span data-ttu-id="ba4df-228">Обычно это делается в методе, как в следующем примере, который назначает обработчики (созданные на более позднем этапе) к событиям и событиям всех диаграмм в `Office.onReady` `onActivated` `onDeactivated` таблице.</span><span class="sxs-lookup"><span data-stu-id="ba4df-228">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="ba4df-229">Далее определите обработчики.</span><span class="sxs-lookup"><span data-stu-id="ba4df-229">Next, define the handlers.</span></span> <span data-ttu-id="ba4df-230">Ниже приводится простой пример ошибки `showDataTab` [HostRestartNeeded,](#handle-the-hostrestartneeded-error) но см. ниже в этой статье для более надежной версии функции.</span><span class="sxs-lookup"><span data-stu-id="ba4df-230">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="ba4df-231">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="ba4df-231">About this code, note:</span></span>

- <span data-ttu-id="ba4df-232">Office определяет время обновления состояния ленты.</span><span class="sxs-lookup"><span data-stu-id="ba4df-232">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="ba4df-233">Метод [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) очереди запроса на обновление.</span><span class="sxs-lookup"><span data-stu-id="ba4df-233">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="ba4df-234">Метод разрешит объект сразу после очереди запроса, а не после обновления `Promise` ленты.</span><span class="sxs-lookup"><span data-stu-id="ba4df-234">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="ba4df-235">Параметром метода является объект `requestUpdate` [RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) указывает вкладку по своему ID точно так, как указано в *JSON* и (2) указывает видимость вкладки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-235">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="ba4df-236">Если у вас есть несколько пользовательских контекстных вкладок, которые должны быть видны в том же контексте, вы просто добавляете дополнительные объекты вкладок в `tabs` массив.</span><span class="sxs-lookup"><span data-stu-id="ba4df-236">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="ba4df-237">Обработник для сокрытия вкладки почти идентичен, за исключением того, что он задает `visible` свойство обратно `false` .</span><span class="sxs-lookup"><span data-stu-id="ba4df-237">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="ba4df-238">Библиотека Office JavaScript также предоставляет несколько интерфейсов (типов), чтобы упростить построение `RibbonUpdateData` объекта.</span><span class="sxs-lookup"><span data-stu-id="ba4df-238">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="ba4df-239">Ниже приводится `showDataTab` функция TypeScript, которая использует эти типы.</span><span class="sxs-lookup"><span data-stu-id="ba4df-239">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="ba4df-240">Обзор вкладок и состояние включенной кнопки одновременно</span><span class="sxs-lookup"><span data-stu-id="ba4df-240">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="ba4df-241">Метод также используется для настройки включенного или отключенного состояния настраиваемой кнопки на настраиваемой контекстной вкладке или настраиваемой основной `requestUpdate` вкладке. Дополнительные сведения см. в материале [Enable and Disable Add-in Commands.](disable-add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="ba4df-241">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="ba4df-242">Возможны сценарии, в которых одновременно необходимо изменить видимость вкладки и состояние включенной кнопки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-242">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="ba4df-243">Это можно сделать одним вызовом `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="ba4df-243">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="ba4df-244">Ниже приводится пример, в котором кнопка на основной вкладке включена одновременно с тем, как отображается контекстная вкладка.</span><span class="sxs-lookup"><span data-stu-id="ba4df-244">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="ba4df-245">В следующем примере включенная кнопка находится на той же контекстной вкладке, которая делается видимой.</span><span class="sxs-lookup"><span data-stu-id="ba4df-245">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="ba4df-246">Локализация BLOB JSON</span><span class="sxs-lookup"><span data-stu-id="ba4df-246">Localizing the JSON blob</span></span>

<span data-ttu-id="ba4df-247">BLOB JSON, который передается, не локализован так же, как локализована разметка манифеста для настраиваемой вкладки ядра (которая описывается при локализации Control из `requestCreateControls` [манифеста).](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="ba4df-247">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="ba4df-248">Вместо этого локализация должна происходить во время запуска с использованием отдельных BLOB-меток JSON для каждого локального.</span><span class="sxs-lookup"><span data-stu-id="ba4df-248">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="ba4df-249">Мы рекомендуем использовать заявление, которое проверяет `switch` [свойство Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage)</span><span class="sxs-lookup"><span data-stu-id="ba4df-249">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="ba4df-250">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="ba4df-250">The following is an example:</span></span>

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

<span data-ttu-id="ba4df-251">Затем код вызывает функцию, чтобы получить локализованный blob, который `requestCreateControls` передается, как в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="ba4df-251">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="ba4df-252">Лучшие практики для настраиваемой контекстной вкладки</span><span class="sxs-lookup"><span data-stu-id="ba4df-252">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="ba4df-253">Реализация альтернативного интерфейса, когда пользовательские контекстные вкладки не поддерживаются</span><span class="sxs-lookup"><span data-stu-id="ba4df-253">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="ba4df-254">Некоторые сочетания платформы, Office приложения и Office сборки не `requestCreateControls` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="ba4df-254">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="ba4df-255">Надстройка должна быть разработана для предоставления альтернативного опыта пользователям, которые запускают надстройки в одной из этих комбинаций.</span><span class="sxs-lookup"><span data-stu-id="ba4df-255">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="ba4df-256">В следующих разделах описаны два способа предоставления впечатления от отката.</span><span class="sxs-lookup"><span data-stu-id="ba4df-256">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="ba4df-257">Использование неконтекстуальных вкладок или элементов управления</span><span class="sxs-lookup"><span data-stu-id="ba4df-257">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="ba4df-258">Существует элемент манифеста [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)который предназначен для создания впечатления от отката в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка запущена на приложении или платформе, которая не поддерживает настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-258">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="ba4df-259">Простейшая стратегия использования этого элемента заключается в *том,* что вы определяете в манифесте одну или несколько настраиваемых вкладки ядра (то есть неконтекстуальные пользовательские вкладки), дублирующие настройки ленты пользовательских контекстных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="ba4df-259">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="ba4df-260">Но вы `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` добавляете в качестве первого детского элемента [CustomTab](../reference/manifest/customtab.md).</span><span class="sxs-lookup"><span data-stu-id="ba4df-260">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="ba4df-261">Эффект от этого ниже:</span><span class="sxs-lookup"><span data-stu-id="ba4df-261">The effect of doing so is the following:</span></span>

- <span data-ttu-id="ba4df-262">Если надстройка работает на приложении и платформе, поддерживаюх настраиваемые контекстные вкладки, то настраиваемая вкладка ядра не будет отображаться на ленте.</span><span class="sxs-lookup"><span data-stu-id="ba4df-262">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="ba4df-263">Вместо этого настраиваемая контекстная вкладка будет создана, когда надстройка вызывает `requestCreateControls` метод.</span><span class="sxs-lookup"><span data-stu-id="ba4df-263">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="ba4df-264">Если надстройка запускается на  приложении или платформе, которые не поддерживаются, на ленте появится настраиваемая вкладка `requestCreateControls` ядра.</span><span class="sxs-lookup"><span data-stu-id="ba4df-264">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="ba4df-265">Ниже приводится пример этой простой стратегии.</span><span class="sxs-lookup"><span data-stu-id="ba4df-265">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="ba4df-266">Эта простая стратегия использует настраиваемую вкладку ядра, которая зеркально отражает настраиваемую контекстную вкладку с ее детскими группами и средствами управления, но можно использовать более сложную стратегию.</span><span class="sxs-lookup"><span data-stu-id="ba4df-266">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="ba4df-267">Элемент также может быть добавлен как (первый) детский элемент к элементам Group и Control (как тип кнопки, так и тип меню), а также `<OverriddenByRibbonApi>` элементам [](../reference/manifest/control.md#button-control) [](../reference/manifest/group.md) [](../reference/manifest/control.md) [](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` меню.</span><span class="sxs-lookup"><span data-stu-id="ba4df-267">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="ba4df-268">Этот факт позволяет распространять группы и элементы управления, которые в противном случае отображаются на контекстной вкладке между различными группами, кнопками и меню в различных настраиваемой основной вкладке.</span><span class="sxs-lookup"><span data-stu-id="ba4df-268">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="ba4df-269">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="ba4df-269">The following is an example.</span></span> <span data-ttu-id="ba4df-270">Обратите внимание, что "MyButton" появится на настраиваемой вкладке ядра только в том случае, если пользовательские контекстные вкладки не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="ba4df-270">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="ba4df-271">Но родительская группа и настраиваемая вкладка ядра будут отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="ba4df-271">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="ba4df-272">Дополнительные примеры см. в [примере OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="ba4df-272">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="ba4df-273">Если родительская вкладка, группа или меню помечены, то она не отображается, и все это детская разметка игнорируется, когда настраиваемые контекстные вкладки не `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="ba4df-273">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="ba4df-274">Поэтому не имеет значения, имеет ли какой-либо из этих детских элементов элемент `<OverriddenByRibbonApi>` или его значение.</span><span class="sxs-lookup"><span data-stu-id="ba4df-274">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="ba4df-275">Следствием этого является то, что если элемент меню, элемент управления или группа должны быть видны во всех контекстах, то не только он не должен быть отмечен, но и его предок меню, группа и вкладка также не должны быть отмечены таким образом `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . </span><span class="sxs-lookup"><span data-stu-id="ba4df-275">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ba4df-276">Не *пометить* все детские элементы вкладки, группы или меню `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="ba4df-276">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="ba4df-277">Это бессмысленно, если родительский элемент помечен по причинам, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` заданным в предыдущем абзаце.</span><span class="sxs-lookup"><span data-stu-id="ba4df-277">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="ba4df-278">Кроме того, если оставить на родительском (или установить его), то родитель будет отображаться независимо от того, поддерживаются ли пользовательские контекстные вкладки, но он будет пустым, когда они `<OverriddenByRibbonApi>` `false` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="ba4df-278">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="ba4df-279">Таким образом, если все элементы ребенка не должны отображаться при поддержке настраиваемой контекстной вкладки, пометите родителя и только родителя с `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="ba4df-279">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="ba4df-280">Использование API, которые показывают или скрывают области задач в указанных контекстах</span><span class="sxs-lookup"><span data-stu-id="ba4df-280">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="ba4df-281">В качестве альтернативы надстройке можно определить области задач с помощью элементов управления пользовательским интерфейсом, дублирующих функции элементов управления на настраиваемой `<OverriddenByRibbonApi>` контекстной вкладке. Затем используйте [методы Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) и [Office.addin.hide,](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) чтобы показать область задач, когда и только когда контекстная вкладка была бы показана, если она была поддержана.</span><span class="sxs-lookup"><span data-stu-id="ba4df-281">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="ba4df-282">Дополнительные сведения об использовании этих методов см. в материале Показать или скрыть области задач [Office надстройки.](../develop/show-hide-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="ba4df-282">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="ba4df-283">Обработка ошибки HostRestartNeeded</span><span class="sxs-lookup"><span data-stu-id="ba4df-283">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="ba4df-284">В некоторых случаях Office не может обновить ленту и возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="ba4df-284">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="ba4df-285">Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть.</span><span class="sxs-lookup"><span data-stu-id="ba4df-285">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="ba4df-286">Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="ba4df-286">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="ba4df-287">Код должен обрабатывать эту ошибку.</span><span class="sxs-lookup"><span data-stu-id="ba4df-287">Your code should handle this error.</span></span> <span data-ttu-id="ba4df-288">Ниже приводится пример того, как.</span><span class="sxs-lookup"><span data-stu-id="ba4df-288">The following is an example of how.</span></span> <span data-ttu-id="ba4df-289">В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.</span><span class="sxs-lookup"><span data-stu-id="ba4df-289">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
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
