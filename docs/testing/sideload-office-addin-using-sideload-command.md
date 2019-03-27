---
title: Загрузка неопубликованных надстроек Office с помощью специальной команды
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: dfa231374133ad857554afaf343362f1415788f4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870116"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="bc60e-102">Загрузка неопубликованных надстроек Office для тестирования с помощью **специальной команды**</span><span class="sxs-lookup"><span data-stu-id="bc60e-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="bc60e-103">Метод "npm run sideload" работает только для надстроек Excel, Word и PowerPoint, которые запускаются в Windows, и только для проектов надстройки, созданных с помощью [инструмента **yo office**](https://github.com/OfficeDev/generator-office), и у которых есть скрипт `sideload` в разделе `scripts` файла package.json.</span><span class="sxs-lookup"><span data-stu-id="bc60e-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="bc60e-104">(У проектов, созданных с помощью более ранних версий **yo office** также нет этого скрипта). Если ваш проект был создан с помощью Visual Studio или у него нет скрипта загрузки неопубликованных приложений, вы можете загрузить его неопубликованным в Windows с помощью метода, описанного в статье [Загрузка неопубликованной надстройки Office из общей сетевой папки](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bc60e-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="bc60e-105">Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей:</span><span class="sxs-lookup"><span data-stu-id="bc60e-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="bc60e-106">Загрузка неопубликованных надстроек Office в Office Online для тестирования</span><span class="sxs-lookup"><span data-stu-id="bc60e-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="bc60e-107">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="bc60e-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="bc60e-108">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="bc60e-108">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="bc60e-109">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="bc60e-109">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="bc60e-110">Измените каталоги на корневой каталог папки вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc60e-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="bc60e-111">Выполните следующую команду, чтобы запустить экземпляр локального сервера на порту 3000 для обслуживания вашего проекта надстройки: "**npm run start**"</span><span class="sxs-lookup"><span data-stu-id="bc60e-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="bc60e-112">Откройте вторую командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="bc60e-112">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="bc60e-113">Измените каталоги на корневой каталог папки вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc60e-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="bc60e-114">Выполните следующую команду для загрузки ведущего приложения (например, Excel или Word) и регистрации надстройки в ведущем приложении: "**npm run sideload**"</span><span class="sxs-lookup"><span data-stu-id="bc60e-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="bc60e-115">См. также</span><span class="sxs-lookup"><span data-stu-id="bc60e-115">See also</span></span>

- [<span data-ttu-id="bc60e-116">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="bc60e-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="bc60e-117">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="bc60e-117">Publish your Office Add-in</span></span>](../publish/publish.md)
