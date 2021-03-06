---
title: Элемент Action в файле манифеста
description: Этот элемент указывает действие, выполняемое, когда пользователь выбирает кнопку или элемент управления меню.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 6be1430800dea27dbd9bf78607161d88e475c145
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505411"
---
# <a name="action-element"></a><span data-ttu-id="54a6c-103">Элемент Action</span><span class="sxs-lookup"><span data-stu-id="54a6c-103">Action element</span></span>

<span data-ttu-id="54a6c-104">Указывает действие, выполняемое при выборе пользователем кнопки [или](control.md#button-control) [управления меню.](control.md#menu-dropdown-button-controls)</span><span class="sxs-lookup"><span data-stu-id="54a6c-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="54a6c-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="54a6c-105">Attributes</span></span>

|  <span data-ttu-id="54a6c-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="54a6c-106">Attribute</span></span>  |  <span data-ttu-id="54a6c-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="54a6c-107">Required</span></span>  |  <span data-ttu-id="54a6c-108">Описание</span><span class="sxs-lookup"><span data-stu-id="54a6c-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="54a6c-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="54a6c-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="54a6c-110">Да</span><span class="sxs-lookup"><span data-stu-id="54a6c-110">Yes</span></span>  | <span data-ttu-id="54a6c-111">Тип выполняемого действия</span><span class="sxs-lookup"><span data-stu-id="54a6c-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="54a6c-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="54a6c-112">Child elements</span></span>

|  <span data-ttu-id="54a6c-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="54a6c-113">Element</span></span> |  <span data-ttu-id="54a6c-114">Описание</span><span class="sxs-lookup"><span data-stu-id="54a6c-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="54a6c-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="54a6c-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="54a6c-116">Указывает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="54a6c-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="54a6c-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="54a6c-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="54a6c-118">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="54a6c-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="54a6c-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="54a6c-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="54a6c-120">Определяет идентификатор контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="54a6c-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="54a6c-121">Title</span><span class="sxs-lookup"><span data-stu-id="54a6c-121">Title</span></span>](#title) | <span data-ttu-id="54a6c-122">Определяет заголовок области задач.</span><span class="sxs-lookup"><span data-stu-id="54a6c-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="54a6c-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="54a6c-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="54a6c-124">Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).</span><span class="sxs-lookup"><span data-stu-id="54a6c-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|

## <a name="xsitype"></a><span data-ttu-id="54a6c-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="54a6c-125">xsi:type</span></span>

<span data-ttu-id="54a6c-p101">Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="54a6c-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> <span data-ttu-id="54a6c-128">Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна при **xsi:type.** `ExecuteFunction`</span><span class="sxs-lookup"><span data-stu-id="54a6c-128">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.</span></span>

## <a name="functionname"></a><span data-ttu-id="54a6c-129">FunctionName</span><span class="sxs-lookup"><span data-stu-id="54a6c-129">FunctionName</span></span>

<span data-ttu-id="54a6c-p102">Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="54a6c-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="54a6c-133">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="54a6c-133">SourceLocation</span></span>

<span data-ttu-id="54a6c-134">Необходимый **элемент, когда xsi:type** — "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="54a6c-134">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="54a6c-135">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="54a6c-135">Specifies the source file location for this action.</span></span> <span data-ttu-id="54a6c-136">Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **URL** в элементе **URL-адресов** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="54a6c-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="54a6c-137">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="54a6c-137">TaskpaneId</span></span>

<span data-ttu-id="54a6c-p104">Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane. Определяет идентификатор для контейнера области задач. Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но оглавление области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.</span><span class="sxs-lookup"><span data-stu-id="54a6c-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="54a6c-143">Этот элемент не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="54a6c-143">This element is not supported in Outlook.</span></span>

<span data-ttu-id="54a6c-144">В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="54a6c-144">The following example shows two actions that share the same **TaskpaneId**.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

<span data-ttu-id="54a6c-p105">В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="54a6c-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="54a6c-147">Должность</span><span class="sxs-lookup"><span data-stu-id="54a6c-147">Title</span></span>

<span data-ttu-id="54a6c-148">Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="54a6c-148">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="54a6c-149">Определяет заголовок области задач для этого действия.</span><span class="sxs-lookup"><span data-stu-id="54a6c-149">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="54a6c-150">В следующем примере показано действие, использующее **элемент Title.**</span><span class="sxs-lookup"><span data-stu-id="54a6c-150">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="54a6c-151">Обратите внимание, что вы не назначаете **заголовок** строке напрямую.</span><span class="sxs-lookup"><span data-stu-id="54a6c-151">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="54a6c-152">Вместо этого вы назначите ему ИД ресурса (resid), который определяется в разделе **Ресурсы** манифеста и может быть не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="54a6c-152">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a><span data-ttu-id="54a6c-153">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="54a6c-153">SupportsPinning</span></span>

<span data-ttu-id="54a6c-154">Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="54a6c-154">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="54a6c-155">Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="54a6c-155">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="54a6c-156">Включите этот элемент со значением `true` для поддержки закрепления области задач.</span><span class="sxs-lookup"><span data-stu-id="54a6c-156">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="54a6c-157">Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента.</span><span class="sxs-lookup"><span data-stu-id="54a6c-157">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="54a6c-158">Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="54a6c-158">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="54a6c-159">Хотя элемент `SupportsPinning` был представлен в наборе [требований 1.5,](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)в настоящее время он поддерживается только для подписчиков Microsoft 365 с помощью следующих элементов.</span><span class="sxs-lookup"><span data-stu-id="54a6c-159">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
>
> - <span data-ttu-id="54a6c-160">Outlook 2016 или более поздней версии Windows (сборка 7628.1000 или более поздней версии)</span><span class="sxs-lookup"><span data-stu-id="54a6c-160">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="54a6c-161">Outlook 2016 или более поздний на Mac (сборка 16.13.503 или более поздней)</span><span class="sxs-lookup"><span data-stu-id="54a6c-161">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="54a6c-162">Современная версия Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="54a6c-162">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
