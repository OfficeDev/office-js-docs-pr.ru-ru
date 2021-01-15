---
title: Элемент Action в файле манифеста
description: Этот элемент указывает действие, выполняемое при выборе пользователем кнопки или элемента управления меню.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: e345d0a1682e0125373a309e1e56eb2d6298ac7d
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771416"
---
# <a name="action-element"></a><span data-ttu-id="9d678-103">Элемент Action</span><span class="sxs-lookup"><span data-stu-id="9d678-103">Action element</span></span>

<span data-ttu-id="9d678-104">Указывает действие, выполняемое при выборе пользователем кнопки [или](control.md#button-control) [меню.](control.md#menu-dropdown-button-controls)</span><span class="sxs-lookup"><span data-stu-id="9d678-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="9d678-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9d678-105">Attributes</span></span>

|  <span data-ttu-id="9d678-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="9d678-106">Attribute</span></span>  |  <span data-ttu-id="9d678-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9d678-107">Required</span></span>  |  <span data-ttu-id="9d678-108">Описание</span><span class="sxs-lookup"><span data-stu-id="9d678-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9d678-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9d678-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9d678-110">Да</span><span class="sxs-lookup"><span data-stu-id="9d678-110">Yes</span></span>  | <span data-ttu-id="9d678-111">Тип выполняемого действия</span><span class="sxs-lookup"><span data-stu-id="9d678-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="9d678-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9d678-112">Child elements</span></span>

|  <span data-ttu-id="9d678-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="9d678-113">Element</span></span> |  <span data-ttu-id="9d678-114">Описание</span><span class="sxs-lookup"><span data-stu-id="9d678-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9d678-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="9d678-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="9d678-116">Указывает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="9d678-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="9d678-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9d678-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="9d678-118">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="9d678-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="9d678-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="9d678-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="9d678-120">Определяет идентификатор контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="9d678-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="9d678-121">Title</span><span class="sxs-lookup"><span data-stu-id="9d678-121">Title</span></span>](#title) | <span data-ttu-id="9d678-122">Определяет заголовок области задач.</span><span class="sxs-lookup"><span data-stu-id="9d678-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="9d678-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="9d678-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="9d678-124">Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).</span><span class="sxs-lookup"><span data-stu-id="9d678-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="9d678-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9d678-125">xsi:type</span></span>

<span data-ttu-id="9d678-p101">Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="9d678-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="9d678-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="9d678-128">FunctionName</span></span>

<span data-ttu-id="9d678-p102">Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="9d678-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="9d678-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9d678-132">SourceLocation</span></span>

<span data-ttu-id="9d678-133">Требуемого **элемента, когда xsi:type** имеет вид "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="9d678-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9d678-134">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="9d678-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="9d678-135">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **Url** в **элементе Urls** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="9d678-135">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="9d678-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="9d678-136">TaskpaneId</span></span>

<span data-ttu-id="9d678-p104">Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane. Определяет идентификатор для контейнера области задач. Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но оглавление области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.</span><span class="sxs-lookup"><span data-stu-id="9d678-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="9d678-142">Этот элемент не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d678-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="9d678-143">В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="9d678-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="9d678-p105">В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="9d678-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="9d678-146">Должность</span><span class="sxs-lookup"><span data-stu-id="9d678-146">Title</span></span>

<span data-ttu-id="9d678-147">Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="9d678-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9d678-148">Определяет заголовок области задач для этого действия.</span><span class="sxs-lookup"><span data-stu-id="9d678-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="9d678-149">В следующем примере показано действие, использующее **элемент Title.**</span><span class="sxs-lookup"><span data-stu-id="9d678-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="9d678-150">Обратите внимание, что заголовок не назначается **строке** напрямую.</span><span class="sxs-lookup"><span data-stu-id="9d678-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="9d678-151">Вместо этого назначьте ему ид ресурса (resid), определенный в разделе **"Ресурсы"** манифеста и не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="9d678-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="9d678-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="9d678-152">SupportsPinning</span></span>

<span data-ttu-id="9d678-153">Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="9d678-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9d678-154">Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="9d678-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="9d678-155">Включите этот элемент со значением `true` для поддержки закрепления области задач.</span><span class="sxs-lookup"><span data-stu-id="9d678-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="9d678-156">Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента.</span><span class="sxs-lookup"><span data-stu-id="9d678-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="9d678-157">Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="9d678-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9d678-158">Хотя элемент был впервые представлен в наборе требований 1.5, в настоящее время он поддерживается только для подписчиков `SupportsPinning` Microsoft 365, использующих следующие [](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)следующую следующую поддержку.</span><span class="sxs-lookup"><span data-stu-id="9d678-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="9d678-159">Outlook 2016 или более поздней версии для Windows (сборка 7628.1000 или более поздней версии)</span><span class="sxs-lookup"><span data-stu-id="9d678-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="9d678-160">Outlook 2016 или более поздней сборки для Mac (сборка 16.13.503 или более поздней)</span><span class="sxs-lookup"><span data-stu-id="9d678-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="9d678-161">Современная версия Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="9d678-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
