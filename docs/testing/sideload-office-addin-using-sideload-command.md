---
title: Загрузка неопубликованных надстроек Office с помощью специальной команды
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 38aa74963ca750d65e4be7bb17745a59eeed0c83
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126892"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="c45b8-102">Загрузка неопубликованных надстроек Office для тестирования с помощью специальной команды</span><span class="sxs-lookup"><span data-stu-id="c45b8-102">Sideload Office Add-ins for testing using the sideload command</span></span>
 
> [!NOTE]
> <span data-ttu-id="c45b8-103">Метод загрузки неопубликованных надстроек, описанный в этой статье, действителен только для:</span><span class="sxs-lookup"><span data-stu-id="c45b8-103">The sideloading technique described in this article is only valid for:</span></span>
> 
> - <span data-ttu-id="c45b8-104">надстроек Excel, Word и PowerPoint, которые выполняются на Windows;</span><span class="sxs-lookup"><span data-stu-id="c45b8-104">Excel, Word, and PowerPoint add-ins that run on Windows</span></span>
> 
> - <span data-ttu-id="c45b8-105">проектов надстроек, которые были созданы с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office) и в которых есть сценарий `sideload` в разделе `scripts` файла package.json.</span><span class="sxs-lookup"><span data-stu-id="c45b8-105">Add-in projects that were created with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="c45b8-106">(В проектах, созданных с помощью более ранних версий генератора Yeoman для надстроек Office, этого сценария не будет.)</span><span class="sxs-lookup"><span data-stu-id="c45b8-106">(Projects that were created with older versions of the Yeoman generator for Office Add-ins will not have this script.)</span></span>
 
<span data-ttu-id="c45b8-107">Для загрузки неопубликованной надстройки с помощью сценария `sideload`, предоставленного генератором Yeoman для надстроек Office, выполните указанные ниже действия:</span><span class="sxs-lookup"><span data-stu-id="c45b8-107">To sideload your add-in by using the `sideload` script that the Yeoman generator for Office Add-ins provides, complete the following steps:</span></span>

1. <span data-ttu-id="c45b8-108">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="c45b8-108">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="c45b8-109">Измените каталоги на корневой каталог папки вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="c45b8-109">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="c45b8-110">Выполните следующую команду, чтобы запустить экземпляр локального веб-сервера на порту 3000 для обслуживания вашего проекта надстройки: "`npm run start`".</span><span class="sxs-lookup"><span data-stu-id="c45b8-110">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: `npm run start`</span></span>

4. <span data-ttu-id="c45b8-111">Откройте вторую командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="c45b8-111">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="c45b8-112">Измените каталоги на корневой каталог папки вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="c45b8-112">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="c45b8-113">Выполните следующую команду для загрузки ведущего приложения (например, Excel или Word) и регистрации надстройки в ведущем приложении: "`npm run sideload`".</span><span class="sxs-lookup"><span data-stu-id="c45b8-113">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: `npm run sideload`</span></span>

<span data-ttu-id="c45b8-114">Если ваш проект надстройки был создан с помощью Visual Studio или у него нет скрипта загрузки неопубликованных приложений, вы можете загрузить его неопубликованным в Windows с помощью метода, описанного в статье [Загрузка неопубликованной надстройки Office из общей сетевой папки](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="c45b8-114">If your add-in project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows by using the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="c45b8-115">Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей для получения сведений о загрузке неопубликованной надстройки:</span><span class="sxs-lookup"><span data-stu-id="c45b8-115">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics for information about sideloading your add-in:</span></span>
 
- [<span data-ttu-id="c45b8-116">Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования</span><span class="sxs-lookup"><span data-stu-id="c45b8-116">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="c45b8-117">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="c45b8-117">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="c45b8-118">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="c45b8-118">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a><span data-ttu-id="c45b8-119">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="c45b8-119">See also</span></span>

- [<span data-ttu-id="c45b8-120">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="c45b8-120">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="c45b8-121">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="c45b8-121">Publish your Office Add-in</span></span>](../publish/publish.md)
