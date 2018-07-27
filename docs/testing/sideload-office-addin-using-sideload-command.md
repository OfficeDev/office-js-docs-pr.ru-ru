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
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Загрузка неопубликованных надстроек Office для тестирования с использованием **команды sideload**
 >[!NOTE]
>(Метод «npm run sideload» работает только для надстроек Excel, Word и PowerPoint.)

1. Откройте командную строку от имени администратора.

2. Измените каталоги на корень папки вашего проекта надстройки.

3. Выполните следующую команду, чтобы запустить экземпляр локального сервера в порт 3000 для подачи вашего проекта надстройки: «**npm run start**»

4. Откройте командную строку от имени администратора.

5. Измените каталоги на корневую папку вашего проекта надстройки.

6. Выполните следующую команду для загрузки ведущего приложения (например, Excel, Word) и регистрации надстройки в ведущем приложении: «**npm run sideload**»

## <a name="see-also"></a>См. также

- [Проверка манифеста и устранение связанных с ним неполадок](troubleshoot-manifest.md)
- [Публикация надстройки Office](../publish/publish.md)