---
title: Краткое руководство по единому входу (SSO)
description: Создание надстройки Office на платформе Node.js с использованием единого входа с помощью генератора Yeoman.
ms.date: 09/07/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: ecbecfd7e475c224451735c7a864f6de2c230d07
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/30/2022
ms.locfileid: "68239383"
---
# <a name="single-sign-on-sso-quick-start"></a>Краткое руководство по единому входу (SSO)

В этой статье вы будете использовать генератор Yeoman для надстроек Office, чтобы создать надстройку Office для Excel, Outlook, Word или PowerPoint, использующую единый вход (SSO).

> [!NOTE]
> Шаблон единого входа, предоставляемый генератором Yeoman для надстроек Office, работает только на локальном узте и не может быть развернут. Если вы создаете новую надстройку Office с единым входом в рабочую среду, следуйте инструкциям в разделе ["](../develop/create-sso-office-add-ins-nodejs.md)Создание надстройки Office Node.js, использующей единый вход".

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases)).

- Последняя версия [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](../develop/yeoman-generator-overview.md). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- Если вы используете компьютер Mac, на котором не установлено приложение Azure CLI, необходимо установить [Homebrew](https://brew.sh/). Сценарий конфигурации единого входа, который вы запустите во время этого быстрого запуска, будет использовать Homebrew для установки Azure CLI, а затем будет использовать Azure CLI для настройки единого входа в Azure.

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

> [!TIP]
> Генератор Yeoman может создать надстройку Office с поддержкой единого входа для Excel, Outlook, Word или PowerPoint с типом скрипта JavaScript или TypeScript. В приведенных ниже инструкциях указаны `JavaScript` и `Excel`, однако следует выбрать тип сценария и клиентское приложение Office, которое лучше всего подходит для вашего сценария.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project supporting single sign-on (localhost)`
- **Выберите тип сценария:** `JavaScript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** Выберите`Excel`, `Outlook`или `Word``Powerpoint`.

:::image type="content" source="../images/yo-office-sso-excel.png" alt-text="Запросы и ответы для генератора Yeoman в интерфейсе командной строки.":::

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит код для надстройки области задач с использованием единого входа.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="configure-sso"></a>Настройка единого входа

Теперь, когда проект надстройки создан и содержит код, необходимый для упрощения процесса единого входа, выполните следующие действия, чтобы настроить единый вход для надстройки.

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Чтобы настроить единый вход для надстройки, выполните приведенную ниже команду.

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > Эта команда приведет к ошибке, если для клиента настроена двухфакторная проверка подлинности. В этом сценарии необходимо вручную выполнить шаги по регистрации приложений Azure и настройке единого входа, выполнив все действия, описанные в руководстве по созданию надстройки [Office Node.js](../develop/create-sso-office-add-ins-nodejs.md) , использующей единый вход.

3. Откроется окно веб-браузера, в котором вам будет предложено войти в Azure. Войдите в Azure, используя учетные данные администратора Microsoft 365. Эти учетные данные будут использоваться для регистрации нового приложения в Azure и настройки параметров, необходимых для единого входа.

    > [!NOTE]
    > Если на этом этапе для входа в Azure вы используете учетные данные без прав администратора, сценарий `configure-sso` не сможет предоставить согласие администратора для надстройки пользователям в организации. В этом случае единый вход будет недоступен для пользователей надстройки, и им будет предложено выполнить вход.

4. После ввода учетных данных закройте окно браузера и вернитесь к командной строке. В процессе настройки единого входа на консоль будут выводиться сообщения о состоянии. В соответствии с ними, файлы проекта надстройки, созданные генератором Yeoman, автоматически обновляются с учетом данных, необходимых для процесса единого входа.

## <a name="test-your-add-in"></a>Тестирование надстройки

Если вы создали надстройку Excel, Word или PowerPoint, выполните действия, описанные в следующем разделе, чтобы опробовать ее. Если вы создали надстройку Outlook, выполните действия, описанные в разделе [Outlook](#outlook) .

### <a name="excel-word-and-powerpoint"></a>Excel, Word и PowerPoint

Выполните следующие действия, чтобы протестировать надстройку Excel, Word или PowerPoint.

1. Когда процесс настройки единого входа будет завершен, для построения проекта, запуска локального веб-сервера и загрузки своей надстройки в ранее выбранное клиентское приложение Office запустите указанную ниже команду.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Когда excel, Word или PowerPoint откроется при выполнении предыдущей команды, убедитесь, что вы вошли с помощью учетной записи пользователя, которая является членом той же организации Microsoft 365, что и учетная запись администратора Microsoft 365, которая использовалась для подключения к Azure при настройке единого входа на шаге 3 предыдущего [раздела](#configure-sso). Благодаря этому будут созданы соответствующие условия для успешного единого входа.

3. В клиентском приложении Office откройте вкладку **"** Главная", а  затем выберите "Показать область задач", чтобы открыть область задач надстройки.

    :::image type="content" source="../images/excel-quickstart-addin-3b.png" alt-text="Кнопка надстройки Excel.":::

4. В нижней части области задач нажмите кнопку **Получить сведения о моем профиле пользователя**, чтобы начать процесс единого входа.

5. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не предоставил надстройке согласие на доступ к Microsoft Graph или пользователь не выполнил вход в Office с помощью действительной учетной записи Майкрософт, Microsoft 365 для образования или рабочей учетной записи. Нажмите **кнопку "Принять** ", чтобы продолжить.

    ![Снимок экрана диалогового окна, запрашивающего разрешение, с выделенной кнопкой "Принять".](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

6. Надстройка получает сведения о профиле пользователя, выполнившего вход, и вносит их в документ. На приведенном ниже рисунке показан пример сведений о профиле, внесенных на лист Excel.

    ![Снимок экрана: сведения о профиле пользователя на листе Excel.](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a>Outlook

Выполните следующие действия, чтобы испытать надстройку Outlook.

1. По завершении процесса настройки единого входа выполните следующую команду, чтобы создать проект и запустить локальный веб-сервер.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md). Убедитесь, что вход в Outlook выполнен в качестве участника той же организации Microsoft 365, что и администратор Microsoft 365, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на шаге 3, описанном в [предыдущем разделе](#configure-sso). Благодаря этому будут созданы соответствующие условия для успешного единого входа.

3. В Outlook создайте новое сообщение.

4. В окне создания сообщения нажмите кнопку **"** Показать область задач", чтобы открыть область задач надстройки.

    ![Снимок экрана: выделенная кнопка ленты надстройки в окне создания сообщения Outlook.](../images/outlook-sso-ribbon-button.png)

5. В нижней части области задач нажмите кнопку **Получить сведения о моем профиле пользователя**, чтобы начать процесс единого входа.

6. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не предоставил надстройке согласие на доступ к Microsoft Graph или пользователь не выполнил вход в Office с помощью действительной учетной записи Майкрософт, Microsoft 365 для образования или рабочей учетной записи. Нажмите **кнопку "Принять** ", чтобы продолжить.

    ![Снимок экрана: диалоговое окно, запрашивающее разрешения, с выделенной кнопкой "Принять".](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

7. Надстройка получает сведения о профиле пользователя, выполнившего вход, и вносит их в текст сообщения электронной почты.

    ![Снимок экрана: сведения о профиле пользователя в окне создания сообщения Outlook.](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы успешно создали надстройку области задач, в которой используется единый вход, когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается. Сведения о настройке надстройки для добавления новых функций, требующих другие разрешения, см. в статье [Настройка надстройки Node.js с поддержкой единого входа](sso-quickstart-customize.md).

## <a name="see-also"></a>См. также

- [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md)
- [Настройка надстройки Node.js с поддержкой единого входа](sso-quickstart-customize.md)
- [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md)
- [Устранение ошибок единого входа](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Использование Visual Studio Code для публикации](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)