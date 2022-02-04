---
title: Элемент Page в файле манифеста
description: 'Элемент Page определяет параметры HTML-страниц, которые настраиваемая функция использует в Excel.'
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="page-element"></a>Элемент Page

Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md) 

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|  Элемент  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями. |

## <a name="example"></a>Пример

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
