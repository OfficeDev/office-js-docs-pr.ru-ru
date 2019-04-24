---
title: Элемент Action в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 59df6cce6af1277f365a1dd3cd0b3ef11230804e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450704"
---
# <a name="action-element"></a><span data-ttu-id="bc8e1-102">Элемент Action</span><span class="sxs-lookup"><span data-stu-id="bc8e1-102">Action element</span></span>

<span data-ttu-id="bc8e1-103">Указывает действие, которое необходимо выполнить, когда пользователь выбирает элемент управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="bc8e1-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bc8e1-104">Attributes</span></span>

|  <span data-ttu-id="bc8e1-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="bc8e1-105">Attribute</span></span>  |  <span data-ttu-id="bc8e1-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bc8e1-106">Required</span></span>  |  <span data-ttu-id="bc8e1-107">Описание</span><span class="sxs-lookup"><span data-stu-id="bc8e1-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="bc8e1-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="bc8e1-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="bc8e1-109">Да</span><span class="sxs-lookup"><span data-stu-id="bc8e1-109">Yes</span></span>  | <span data-ttu-id="bc8e1-110">Тип выполняемого действия</span><span class="sxs-lookup"><span data-stu-id="bc8e1-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="bc8e1-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bc8e1-111">Child elements</span></span>

|  <span data-ttu-id="bc8e1-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="bc8e1-112">Element</span></span> |  <span data-ttu-id="bc8e1-113">Описание</span><span class="sxs-lookup"><span data-stu-id="bc8e1-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="bc8e1-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="bc8e1-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="bc8e1-115">Указывает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="bc8e1-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bc8e1-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="bc8e1-117">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="bc8e1-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="bc8e1-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="bc8e1-119">Определяет идентификатор для контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="bc8e1-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="bc8e1-120"> [Title](#title)</span></span> | <span data-ttu-id="bc8e1-121">Определяет заголовок области задач.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="bc8e1-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="bc8e1-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="bc8e1-123">Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="bc8e1-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="bc8e1-124">xsi:type</span></span>

<span data-ttu-id="bc8e1-p101">Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="bc8e1-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="bc8e1-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="bc8e1-127">FunctionName</span></span>

<span data-ttu-id="bc8e1-p102">Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="bc8e1-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bc8e1-131">SourceLocation</span></span>

<span data-ttu-id="bc8e1-p103">Обязательный элемент, если атрибуту **xsi:type** присвоено значение ShowTaskpane. Указывает расположение исходного файла для этого действия. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **Url** в элементе **Urls**, включенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="bc8e1-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="bc8e1-135">TaskpaneId</span></span>

<span data-ttu-id="bc8e1-136">Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="bc8e1-137">Определяет идентификатор контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="bc8e1-138">Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="bc8e1-139">Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="bc8e1-140">Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но содержимое области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="bc8e1-141">Этот элемент не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="bc8e1-142">В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="bc8e1-p105">В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="bc8e1-145">Должность</span><span class="sxs-lookup"><span data-stu-id="bc8e1-145">Title</span></span>

<span data-ttu-id="bc8e1-146">Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="bc8e1-147">Определяет заголовок области задач для этого действия.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="bc8e1-148">В приведенных ниже примерах показаны два действия, для которых используется элемент **Title**.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-148">The following examples show two different actions that use the **Title** element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a><span data-ttu-id="bc8e1-149">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="bc8e1-149">SupportsPinning</span></span>

<span data-ttu-id="bc8e1-150">Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-150">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="bc8e1-151">Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-151">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="bc8e1-152">Включите этот элемент со значением `true` для поддержки закрепления области задач.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-152">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="bc8e1-153">Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента.</span><span class="sxs-lookup"><span data-stu-id="bc8e1-153">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="bc8e1-154">Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-154">For more information, see [Implement a pinnable task pane in Outlook](/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="bc8e1-155">Суппортспиннинг в настоящее время поддерживается только Outlook 2016 для Windows (сборка 7628,1000 или более поздней версии) и Outlook 2016 для Mac (сборка 16.13.503 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="bc8e1-155">SupportsPinning is currently only supported by Outlook 2016 for Windows (build 7628.1000 or later) and Outlook 2016 for Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
