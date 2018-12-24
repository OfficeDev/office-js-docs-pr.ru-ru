---
title: Элемент Permissions в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 9193651ec0c795cdb55eb3fc6576dbacd59e0fb2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432356"
---
# <a name="permissions-element"></a><span data-ttu-id="aacbc-102">Элемент Permissions</span><span class="sxs-lookup"><span data-stu-id="aacbc-102">Permissions element</span></span>

<span data-ttu-id="aacbc-103">Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.</span><span class="sxs-lookup"><span data-stu-id="aacbc-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="aacbc-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="aacbc-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="aacbc-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="aacbc-105">Syntax</span></span>

<span data-ttu-id="aacbc-106">Для надстроек области задач и контентных надстроек:</span><span class="sxs-lookup"><span data-stu-id="aacbc-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="aacbc-107">Для почтовых надстроек</span><span class="sxs-lookup"><span data-stu-id="aacbc-107">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="aacbc-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="aacbc-108">Contained in</span></span>

[<span data-ttu-id="aacbc-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="aacbc-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="aacbc-110">Примечания</span><span class="sxs-lookup"><span data-stu-id="aacbc-110">Remarks</span></span>

<span data-ttu-id="aacbc-111">Подробные сведения см. в статьях [Запрашивание разрешений на использование API в надстройках области задач и контентных надстройках](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) и [Общие сведения о разрешениях для надстроек Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="aacbc-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
