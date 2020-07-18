---
title: Тестирование и отладка надстроек Office
description: Узнайте, как тестировать и отлаживать свою надстройку Office
ms.date: 06/17/2020
localization_priority: Priority
ms.openlocfilehash: 526204fe94d4c97ce7e1e0bc9ac2a212f69611d3
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159250"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="398f6-103">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="398f6-103">Test and debug Office Add-ins</span></span>

<span data-ttu-id="398f6-104">Этот раздел содержит рекомендации по тестированию, отладке и диагностике надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="398f6-104">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="398f6-105">Загрузка неопубликованной надстройки Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="398f6-105">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="398f6-p101">Вы можете установить надстройку Office для тестирования, не размещая ее в каталоге надстроек. Процедура отличается для разных платформ, а в некоторых случаях и для разных продуктов. Следующие статьи посвящены загрузке неопубликованных надстроек Office на определенной платформе или в определенном продукте:</span><span class="sxs-lookup"><span data-stu-id="398f6-p101">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="398f6-109">Загрузка неопубликованных надстроек Office в Windows</span><span class="sxs-lookup"><span data-stu-id="398f6-109">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="398f6-110">Загрузка неопубликованных надстроек Office в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="398f6-110">Sideload Office Add-ins in Office on the web</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="398f6-111">Загрузка неопубликованных надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="398f6-111">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="398f6-112">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="398f6-112">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="398f6-113">Отладка надстройки Office</span><span class="sxs-lookup"><span data-stu-id="398f6-113">Debug an Office Add-in</span></span>

<span data-ttu-id="398f6-p102">Процедура отладки также отличается для разных платформ. Следующие статьи посвящены отладке надстроек Office на определенной платформе:</span><span class="sxs-lookup"><span data-stu-id="398f6-p102">The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="398f6-116">Подключение отладчика из области задач (в Windows)</span><span class="sxs-lookup"><span data-stu-id="398f6-116">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="398f6-117">Отладка надстроек с помощью средств разработчика F12 в Windows 10</span><span class="sxs-lookup"><span data-stu-id="398f6-117">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="398f6-118">Отладка надстроек в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="398f6-118">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="398f6-119">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="398f6-119">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

- [<span data-ttu-id="398f6-120">Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"</span><span class="sxs-lookup"><span data-stu-id="398f6-120">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="398f6-121">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="398f6-121">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="398f6-122">Информацию о проверке манифеста надстройки Office и устранении связанных с ним неполадок см. в [этой статье](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="398f6-122">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="398f6-123">Устранение ошибок, с которыми сталкиваются пользователи</span><span class="sxs-lookup"><span data-stu-id="398f6-123">Troubleshoot user errors</span></span>

<span data-ttu-id="398f6-124">Информацию об устранении основных ошибок, с которыми сталкиваются пользователи при работе с надстройками Office, см. в [этой статье](testing-and-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="398f6-124">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
