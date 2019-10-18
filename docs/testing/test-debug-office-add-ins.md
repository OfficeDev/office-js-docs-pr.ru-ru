---
title: Тестирование и отладка надстроек Office
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 3c841608d36f5004a876bec2c899d0e5659d47a7
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126920"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="542b5-102">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="542b5-102">Test and debug Office Add-ins</span></span>

<span data-ttu-id="542b5-103">Этот раздел содержит рекомендации по тестированию, отладке и диагностике надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="542b5-103">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="542b5-104">Загрузка неопубликованной надстройки Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="542b5-104">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="542b5-p101">Вы можете установить надстройку Office для тестирования, не размещая ее в каталоге надстроек. Процедура отличается для разных платформ, а в некоторых случаях и для разных продуктов. Следующие статьи посвящены загрузке неопубликованных надстроек Office на определенной платформе или в определенном продукте:</span><span class="sxs-lookup"><span data-stu-id="542b5-p101">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="542b5-108">Загрузка неопубликованных надстроек Office в Windows</span><span class="sxs-lookup"><span data-stu-id="542b5-108">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="542b5-109">Загрузка неопубликованных надстроек Office в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="542b5-109">Sideload Office Add-ins in Office on the web</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="542b5-110">Загрузка неопубликованных надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="542b5-110">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="542b5-111">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="542b5-111">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="542b5-112">Отладка надстройки Office</span><span class="sxs-lookup"><span data-stu-id="542b5-112">Debug an Office Add-in</span></span>

<span data-ttu-id="542b5-p102">Процедура отладки также отличается для разных платформ. Следующие статьи посвящены отладке надстроек Office на определенной платформе:</span><span class="sxs-lookup"><span data-stu-id="542b5-p102">The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="542b5-115">Подключение отладчика из области задач (в Windows)</span><span class="sxs-lookup"><span data-stu-id="542b5-115">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="542b5-116">Отладка надстроек с помощью средств разработчика F12 в Windows 10</span><span class="sxs-lookup"><span data-stu-id="542b5-116">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="542b5-117">Отладка надстроек в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="542b5-117">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="542b5-118">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="542b5-118">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="542b5-119">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="542b5-119">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="542b5-120">Информацию о проверке манифеста надстройки Office и устранении связанных с ним неполадок см. в [этой статье](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="542b5-120">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="542b5-121">Устранение ошибок, с которыми сталкиваются пользователи</span><span class="sxs-lookup"><span data-stu-id="542b5-121">Troubleshoot user errors</span></span>

<span data-ttu-id="542b5-122">Информацию об устранении основных ошибок, с которыми сталкиваются пользователи при работе с надстройками Office, см. в [этой статье](testing-and-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="542b5-122">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
