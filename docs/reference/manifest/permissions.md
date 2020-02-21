---
title: Элемент Permissions в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 95cb45f89e2a5b92edc29bf32d0b47fcb2dbf8ce
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165547"
---
# <a name="permissions-element"></a>Элемент Permissions

Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

Для надстроек области задач и контентных надстроек:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Для почтовых надстроек

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Более подробную информацию можно узнать [в статье запрос разрешений для использования API в](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) надстройках и [Общие сведения о разрешениях для надстроек Outlook](../../outlook/understanding-outlook-add-in-permissions.md).
