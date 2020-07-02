---
title: Элемент Permissions в файле манифеста
description: Элемент Permissions определяет уровень доступа к API для надстройки Office.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: bc4cc2713d5a781c3407385470acd762910d17fd
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006460"
---
# <a name="permissions-element"></a>Элемент Permissions

Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

Для надстроек области задач и контентных надстроек:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Для почтовых надстроек:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Дополнительные сведения см в статье [запрашивание разрешений для использования API в контентных](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) надстройках и надстройках области задач и [Общие сведения о разрешениях для надстроек Outlook](../../outlook/understanding-outlook-add-in-permissions.md).
