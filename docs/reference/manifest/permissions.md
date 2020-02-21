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
# <a name="permissions-element"></a><span data-ttu-id="2f57c-102">Элемент Permissions</span><span class="sxs-lookup"><span data-stu-id="2f57c-102">Permissions element</span></span>

<span data-ttu-id="2f57c-103">Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.</span><span class="sxs-lookup"><span data-stu-id="2f57c-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="2f57c-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2f57c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2f57c-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2f57c-105">Syntax</span></span>

<span data-ttu-id="2f57c-106">Для надстроек области задач и контентных надстроек:</span><span class="sxs-lookup"><span data-stu-id="2f57c-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="2f57c-107">Для почтовых надстроек</span><span class="sxs-lookup"><span data-stu-id="2f57c-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="2f57c-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2f57c-108">Contained in</span></span>

[<span data-ttu-id="2f57c-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2f57c-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="2f57c-110">Замечания</span><span class="sxs-lookup"><span data-stu-id="2f57c-110">Remarks</span></span>

<span data-ttu-id="2f57c-111">Более подробную информацию можно узнать [в статье запрос разрешений для использования API в](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) надстройках и [Общие сведения о разрешениях для надстроек Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="2f57c-111">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
