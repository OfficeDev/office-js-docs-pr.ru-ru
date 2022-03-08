---
title: Элемент Action в файле манифеста
description: Этот элемент указывает действие, выполняемое, когда пользователь выбирает кнопку или элемент управления меню.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21c8f9a6345641f23aad70efed67c9c45f72a1c8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340416"
---
# <a name="action-element"></a>Элемент Action

Указывает действие, выполняемое при выборе пользователем кнопки  [или](control-button.md) [управления меню](control-menu.md) .

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Тип выполняемого действия|

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Указывает имя выполняемой функции. |
|  [SourceLocation](#sourcelocation) |    Указывает расположение исходного файла для этого действия. |
|  [TaskpaneId](#taskpaneid) | Определяет идентификатор для контейнера области задач. Не поддерживается Outlook надстройки.|
|  [Title](#title) | Определяет заголовок области задач. Не поддерживается Outlook надстройки.|
|  [SupportsPinning](#supportspinning) | Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).|

## <a name="xsitype"></a>xsi:type

Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна при **xsi:type**`ExecuteFunction`.

## <a name="functionname"></a>FunctionName

Необходимый элемент **при xsi:type** .`ExecuteFunction` Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Необходимый элемент **при xsi:type** .`ShowTaskpane` Указывает расположение исходного файла для этого действия. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **URL** в элементе **URL-адресов** в [элементе Resources](resources.md) .

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Необязательный элемент **при xsi:type** .`ShowTaskpane` Определяет идентификатор для контейнера области задач. Если у вас есть несколько `ShowTaskpane` действий, используйте другой **taskpaneId** , если для каждого из них нужна независимая панорама. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, которые имеют один и тот же **TaskpaneId**, контейнер области останется открытым, но содержимое области будет заменено соответствующим действием `SourceLocation`.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

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

Необязательный элемент **при xsi:type** .`ShowTaskpane` Определяет заголовок области задач для этого действия.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> Этот элемент не поддерживается в Outlook надстройки.

В следующем примере показано действие, использующее **элемент Title** . Обратите внимание, что вы не назначаете **заголовок** строке напрямую. Вместо этого вы назначите ему ИД ресурса (resid), который определяется в разделе **Ресурсы** манифеста и может быть не более 32 символов.

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

Необязательный элемент **при xsi:type** .`ShowTaskpane` Элементы [, содержащие VersionOverrides](versionoverrides.md), должны иметь **значение атрибута xsi:type** .`VersionOverridesV1_1` Включите этот элемент со значением `true` для поддержки закрепления области задач. Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента. Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Mailbox 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

> [!IMPORTANT]
> Хотя элемент **SupportsPinning** был представлен в наборе [требований 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), в настоящее время он поддерживается только для Microsoft 365 абонентов с помощью следующих ниже:
>
> - Outlook 2016 или более поздней Windows (сборка 7628.1000 или более поздней)
> - Outlook 2016 или позже на Mac (сборка 16.13.503 или более поздней сборки)
> - Современная версия Outlook в Интернете

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
