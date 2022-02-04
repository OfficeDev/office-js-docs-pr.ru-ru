---
title: Элемент SourceLocation для настраиваемой функции в файле манифеста
description: 'Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.'
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="sourcelocation-element-custom-functions"></a>Элемент SourceLocation (настраиваемые функции)

Определяет расположение ресурса, необходимого элементам **Script** или **Page**, используемым пользовательскими функциями в Excel.

> [!IMPORTANT]
> Эта статья относится только к **SourceLocation** , которая является ребенком элементов **Page** или **Script** . Сведения о элементе **SourceLocation базового** манифеста см. в [sourceLocation](sourcelocation.md).

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Да      | Имя ресурса URL-адреса, определенного в разделе **Ресурсы** в манифесте. Может быть не более 32 символов. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<SourceLocation resid="pageURL"/>
```
