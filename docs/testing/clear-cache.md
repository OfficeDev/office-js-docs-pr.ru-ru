---
title: Очистка кэша Office
description: Узнайте, как очищать кэш Office на компьютере.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915080"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="9d6a3-103">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="9d6a3-103">Clear the Office cache</span></span>

<span data-ttu-id="9d6a3-104">Можно удалить надстройку, ранее установленную в Windows, на компьютерах Mac или в iOS, очистив кэш Office на компьютере.</span><span class="sxs-lookup"><span data-stu-id="9d6a3-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="9d6a3-105">Кроме того, если вы изменяете манифест надстройки (например, обновляете имена файлов значков или текст команд надстройки), следует очистить кэш Office, а потом заново установить надстройку с помощью обновленного манифеста.</span><span class="sxs-lookup"><span data-stu-id="9d6a3-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="9d6a3-106">В этом случае надстройка будет отображаться в Office в соответствии с обновленным манифестом.</span><span class="sxs-lookup"><span data-stu-id="9d6a3-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="9d6a3-107">Очистка кэша Office в Windows</span><span class="sxs-lookup"><span data-stu-id="9d6a3-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="9d6a3-108">Чтобы очистить кэш Office в Windows, удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="9d6a3-108">To clear the Office cache on Windows, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="9d6a3-109">Очистка кэша Office на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="9d6a3-109">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="9d6a3-110">Очистка кэша Office в iOS</span><span class="sxs-lookup"><span data-stu-id="9d6a3-110">Clear the Office cache on iOS</span></span>

<span data-ttu-id="9d6a3-111">Чтобы очистить кэш Office в iOS, вызовите `window.location.reload(true)` в JavaScript в надстройке, чтобы запустить принудительную перезагрузку.</span><span class="sxs-lookup"><span data-stu-id="9d6a3-111">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="9d6a3-112">Также можно переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="9d6a3-112">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="9d6a3-113">См. также</span><span class="sxs-lookup"><span data-stu-id="9d6a3-113">See also</span></span>

- [<span data-ttu-id="9d6a3-114">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9d6a3-114">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="9d6a3-115">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9d6a3-115">Validate an Office Add-in manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="9d6a3-116">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="9d6a3-116">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="9d6a3-117">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="9d6a3-117">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="9d6a3-118">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9d6a3-118">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)