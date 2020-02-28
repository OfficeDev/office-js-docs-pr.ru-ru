---
title: Элемент Action в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b05da08f4995c7d8f7270e7fba6f416c9903b066
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324893"
---
# <a name="action-element"></a><span data-ttu-id="0b165-102">Элемент Action</span><span class="sxs-lookup"><span data-stu-id="0b165-102">Action element</span></span>

<span data-ttu-id="0b165-103">Указывает действие, которое необходимо выполнить, когда пользователь выбирает элемент управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="0b165-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="0b165-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b165-104">Attributes</span></span>

|  <span data-ttu-id="0b165-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="0b165-105">Attribute</span></span>  |  <span data-ttu-id="0b165-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0b165-106">Required</span></span>  |  <span data-ttu-id="0b165-107">Описание</span><span class="sxs-lookup"><span data-stu-id="0b165-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0b165-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0b165-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="0b165-109">Да</span><span class="sxs-lookup"><span data-stu-id="0b165-109">Yes</span></span>  | <span data-ttu-id="0b165-110">Тип выполняемого действия</span><span class="sxs-lookup"><span data-stu-id="0b165-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="0b165-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="0b165-111">Child elements</span></span>

|  <span data-ttu-id="0b165-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b165-112">Element</span></span> |  <span data-ttu-id="0b165-113">Описание</span><span class="sxs-lookup"><span data-stu-id="0b165-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="0b165-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="0b165-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="0b165-115">Указывает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="0b165-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="0b165-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0b165-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="0b165-117">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="0b165-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="0b165-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="0b165-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="0b165-119">Определяет идентификатор для контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="0b165-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="0b165-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="0b165-120"> [Title](#title)</span></span> | <span data-ttu-id="0b165-121">Определяет заголовок области задач.</span><span class="sxs-lookup"><span data-stu-id="0b165-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="0b165-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="0b165-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="0b165-123">Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).</span><span class="sxs-lookup"><span data-stu-id="0b165-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="0b165-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0b165-124">xsi:type</span></span>

<span data-ttu-id="0b165-p101">Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="0b165-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="0b165-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="0b165-127">FunctionName</span></span>

<span data-ttu-id="0b165-p102">Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="0b165-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="0b165-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0b165-131">SourceLocation</span></span>

<span data-ttu-id="0b165-132">Обязательный элемент, если **xsi: Type** — "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="0b165-132">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0b165-133">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="0b165-133">Specifies the source file location for this action.</span></span> <span data-ttu-id="0b165-134">Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **Url** в элементе **Urls**, включенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0b165-134">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="0b165-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="0b165-135">TaskpaneId</span></span>

<span data-ttu-id="0b165-136">Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="0b165-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0b165-137">Определяет идентификатор контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="0b165-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="0b165-138">Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="0b165-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="0b165-139">Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область.</span><span class="sxs-lookup"><span data-stu-id="0b165-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="0b165-140">Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но содержимое области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.</span><span class="sxs-lookup"><span data-stu-id="0b165-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="0b165-141">Этот элемент не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="0b165-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="0b165-142">В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="0b165-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="0b165-p105">В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="0b165-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="0b165-145">Должность</span><span class="sxs-lookup"><span data-stu-id="0b165-145">Title</span></span>

<span data-ttu-id="0b165-146">Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="0b165-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0b165-147">Определяет заголовок области задач для этого действия.</span><span class="sxs-lookup"><span data-stu-id="0b165-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="0b165-148">В приведенном ниже примере показано действие, в котором используется элемент **Title** .</span><span class="sxs-lookup"><span data-stu-id="0b165-148">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="0b165-149">Обратите внимание, что **заголовок** не назначается строке напрямую.</span><span class="sxs-lookup"><span data-stu-id="0b165-149">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="0b165-150">Вместо этого ему назначается идентификатор ресурса (Resid), который определяется в разделе **Resources (ресурсы** ) манифеста.</span><span class="sxs-lookup"><span data-stu-id="0b165-150">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="0b165-151">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="0b165-151">SupportsPinning</span></span>

<span data-ttu-id="0b165-152">Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="0b165-152">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="0b165-153">Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="0b165-153">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="0b165-154">Включите этот элемент со значением `true` для поддержки закрепления области задач.</span><span class="sxs-lookup"><span data-stu-id="0b165-154">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="0b165-155">Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента.</span><span class="sxs-lookup"><span data-stu-id="0b165-155">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="0b165-156">Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="0b165-156">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!NOTE]
> <span data-ttu-id="0b165-157">Суппортспиннинг в настоящее время поддерживается только в Outlook 2016 или более поздних версий для Windows (сборка 7628,1000 или более поздней версии) и Outlook 2016 или более поздней версии в Mac (сборка 16.13.503 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="0b165-157">SupportsPinning is currently only supported by Outlook 2016 or later on Windows (build 7628.1000 or later) and Outlook 2016 or later on Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
