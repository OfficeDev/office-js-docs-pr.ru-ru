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
# <a name="action-element"></a>Элемент Action

Указывает действие, выполняемое при выборе пользователем кнопки [или](control.md#button-control) [меню.](control.md#menu-dropdown-button-controls)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Тип выполняемого действия|

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Указывает имя выполняемой функции. |
|  [SourceLocation](#sourcelocation) |    Указывает расположение исходного файла для этого действия. |
|  [TaskpaneId](#taskpaneid) | Определяет идентификатор контейнера области задач.|
|  [Title](#title) | Определяет заголовок области задач.|
|  [SupportsPinning](#supportspinning) | Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).|
  

## <a name="xsitype"></a>xsi:type

Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Требуемого **элемента, когда xsi:type** имеет вид "ShowTaskpane". Указывает расположение исходного файла для этого действия. Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **Url** в **элементе Urls** в [элементе Resources.](resources.md)

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane. Определяет идентификатор для контейнера области задач. Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но оглавление области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.

> [!NOTE]
> Этот элемент не поддерживается в Outlook.

В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.

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

В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

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

## <a name="title"></a>Должность

Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane. Определяет заголовок области задач для этого действия.

В следующем примере показано действие, использующее **элемент Title.** Обратите внимание, что заголовок не назначается **строке** напрямую. Вместо этого назначьте ему ид ресурса (resid), определенный в разделе **"Ресурсы"** манифеста и не более 32 символов.

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

## <a name="supportspinning"></a>SupportsPinning

Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane. Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`. Включите этот элемент со значением `true` для поддержки закрепления области задач. Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента. Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).

> [!IMPORTANT]
> Хотя элемент был впервые представлен в наборе требований 1.5, в настоящее время он поддерживается только для подписчиков `SupportsPinning` Microsoft 365, использующих следующие [](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)следующую следующую поддержку.
> - Outlook 2016 или более поздней версии для Windows (сборка 7628.1000 или более поздней версии)
> - Outlook 2016 или более поздней сборки для Mac (сборка 16.13.503 или более поздней)
> - Современная версия Outlook в Интернете

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
