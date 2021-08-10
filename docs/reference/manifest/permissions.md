---
title: Элемент Permissions в файле манифеста
description: Элемент Permissions указывает уровень доступа к API для Office надстройки.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: 2f2ccb4f6ec691b19cadea76a06520a9bad7a0b6c0e51699f2c8db67a3030de0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089016"
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

Дополнительные сведения см. в материале Запрашивание разрешений на использование [API](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) в надстройке контента и области задач и Outlook разрешений [надстройки.](../../outlook/understanding-outlook-add-in-permissions.md)
