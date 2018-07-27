---
title: Загрузка неопубликованных надстроек Office с использованием команды sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279362"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="49b8b-102">Загрузка неопубликованных надстроек Office для тестирования с использованием **команды sideload**</span><span class="sxs-lookup"><span data-stu-id="49b8b-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="49b8b-103">(Метод «npm run sideload» работает только для надстроек Excel, Word и PowerPoint.)</span><span class="sxs-lookup"><span data-stu-id="49b8b-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

1. <span data-ttu-id="49b8b-104">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="49b8b-104">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="49b8b-105">Измените каталоги на корень папки вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="49b8b-105">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="49b8b-106">Выполните следующую команду, чтобы запустить экземпляр локального сервера в порт 3000 для подачи вашего проекта надстройки: «**npm run start**»</span><span class="sxs-lookup"><span data-stu-id="49b8b-106">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="49b8b-107">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="49b8b-107">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="49b8b-108">Измените каталоги на корневую папку вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="49b8b-108">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="49b8b-109">Выполните следующую команду для загрузки ведущего приложения (например, Excel, Word) и регистрации надстройки в ведущем приложении: «**npm run sideload**»</span><span class="sxs-lookup"><span data-stu-id="49b8b-109">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="49b8b-110">См. также</span><span class="sxs-lookup"><span data-stu-id="49b8b-110">See also</span></span>

- [<span data-ttu-id="49b8b-111">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="49b8b-111">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="49b8b-112">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="49b8b-112">Publish your Office Add-in</span></span>](../publish/publish.md)