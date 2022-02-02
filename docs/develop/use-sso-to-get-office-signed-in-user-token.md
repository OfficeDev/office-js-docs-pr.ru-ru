---
title: Использование SSO для получения удостоверения пользователя, вписаного в него.
description: Позвоните в API getAccessToken, чтобы получить маркер ID с именем, электронной почтой и дополнительной информацией о подписанной пользователем.
ms.date: 01/25/2022
localization_priority: Normal
ms.openlocfilehash: 2c9b3c89a154d624f99e196014c7d8024286d927
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/02/2022
ms.locfileid: "62322338"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>Использование SSO для получения удостоверения пользователя, вписаного в него.

Используйте API`getAccessToken`, чтобы получить маркер доступа, содержащий удостоверение для текущего пользователя, вписающегося в Office. Маркер доступа также является маркером идентификатора, так как он содержит утверждения о удостоверениях, подписанных пользователем, таких как их имя и электронная почта. Вы также можете использовать маркер ID для идентификации пользователя при вызове собственных веб-служб. Чтобы позвонить`getAccessToken`, необходимо настроить Office надстройку, чтобы использовать SSO с Office.

В этой статье вы создайте надстройку Office, которая получает маркер ID и отображает имя пользователя, электронную почту и уникальный ID в области задач.

> [!NOTE]
> SSO с Office и `getAccessToken` API не работают во всех сценариях. Всегда реализуйте диалоговое окно отката, чтобы войти в пользователя, когда SSO недоступен. Дополнительные сведения см. в ссылке [Authenticate and authorize with the Office диалоговом API](auth-with-office-dialog-api.md).

## <a name="create-an-app-registration"></a>Создание регистрации приложения

Чтобы использовать SSO с Office, необходимо создать регистрацию приложений на портале Azure, чтобы платформа удостоверений Майкрософт могли предоставлять службы проверки подлинности и авторизации для вашей надстройки Office и ее пользователей.

1. Чтобы зарегистрировать свое приложение, перейдите на портал [Azure — страницу регистрации приложений](https://go.microsoft.com/fwlink/?linkid=2083908) .

1. Вопишитесь с **_учетными_** данными администратора в Microsoft 365 аренды. Пример: MyName@contoso.onmicrosoft.com.

1. Выберите **Новая регистрация**. На странице **Зарегистрировать приложение** задайте необходимые значения следующим образом.

   - Введите **имя** `Office-Add-in-SSO`.
   - Для параметра **Поддерживаемые типы учетных записей** укажите вариант **Учетные записи в любом каталоге организации и личные учетные записи Майкрософт (например, Skype, Xbox, Outlook.com)**.
   - Установите тип приложения в **Интернете,** а затем установите **URI перенаправления**`https://localhost:[port]/dialog.html`. Замените `[port]` правильный номер порта для веб-приложения. Если надстройка создана с помощью yo office, номер порта обычно составляет 3000 и находится в файле package.json. Если надстройка создана с Visual Studio 2019 г., порт находится в **свойстве URL-адреса SSL** веб-проекта.
   - Нажмите кнопку **Зарегистрировать**.

1. На странице **Office-Надстройка-SSO** скопируйте и сохраните значения для **ID приложения (клиента)** и **ID каталога (клиента**). Они понадобятся вам позже.

   > [!NOTE]
   > Этот **ID приложения (клиента)** — это значение "аудитории", когда другие приложения, например клиентская Office (например, PowerPoint, Word, Excel), ищут авторизованный доступ к приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.

1. Выберите **Проверка подлинности** в разделе **Управление**. В разделе **Неявный грант** впускаете почтовые ящики для маркера **Доступа и** **маркера ID**.

1. Щелкните **Сохранить** в верхней части формы.

1. Выберите пункт **Предоставление API** в разделе **Управление**. Выберите **ссылку Set** . Это позволит создать URI приложения iD в форме `api://[app-id-guid]`, `[app-id-guid]` где находится **ID приложения (клиента**).

1. В сгенерированном ID `localhost:[port]/` вставьте (обратите внимание на переназначенную полосу "/" прим. в конце) между двумя полосами вперед и GUID. Замените `[port]` правильный номер порта для веб-приложения. Если надстройка создана с помощью yo office, номер порта обычно составляет 3000 и находится в файле package.json. Если надстройка создана с Visual Studio 2019 г., порт находится в **свойстве URL-адреса SSL** веб-проекта.
   Когда вы закончите, весь ID должен иметь форму `api://localhost:[port]/[app-id-guid]`; например `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Нажмите кнопку **Добавить область**. В открывшейся панели введите `access_as_user` в качестве параметра **Имя области**.

1. Для параметра **Кто может давать согласие?** установите вариант **Администраторы и пользователи**.

1. `access_as_user` Заполните поля для настройки API администратора и согласия пользователя со значениями, подходящими для области, которая позволяет приложению Office использовать веб-API надстройки с тем же правами, что и текущий пользователь. Предложения:

   - **Имя отображения согласия** администратора: Office может выступать в качестве пользователя.
   - **Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у текущего пользователя.
   - **Имя отображения согласия** пользователя: Office может действовать так, как вы.
   - **Описание согласия пользователя**: Office включить вызов веб-API надстройки с тем же правами, что и у вас.

1. Убедитесь, что параметру **Состояние** присвоено значение **Включено**.

1. Нажмите кнопку **Добавить область**.

   > [!NOTE]
   > Доменная часть имени **области**, отображаемая непосредственно под текстовым полем, должна автоматически соответствовать URI идентификатора приложения, заданного ранее, с добавлением `/access_as_user` в конце, например: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. В разделе **Авторизованные клиентские приложения** укажите приложения, которые необходимо авторизовать для веб-приложения надстройки. Необходимо обеспечить предварительную авторизацию для всех указанных ниже идентификаторов.

   - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).
   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).
   - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).
   - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office в Интернете).
   - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook в Интернете).

   Для каждого идентификатора сделайте следующее:

   а. Выберите **кнопку** Добавить кнопку клиентского приложения, а затем в `[app-id-guid]` открываемой панели установите ID приложения (клиента) и проверьте поле `api://localhost:44355/[app-id-guid]/access_as_user`для .

   б. Нажмите кнопку **Добавить приложение**.

1. Выберите пункт **Разрешения API** в разделе **Управление** и нажмите кнопку **Добавить разрешение**. В открывшейся панели выберите **Microsoft Graph** и щелкните **Делегированные разрешения**.

1. Используйте поле поиска **Выбрать разрешения**, чтобы найти нужные разрешения для надстройки. Поиск и выбор **разрешения профиля** . Для `profile` получения маркера в веб-приложении Office требуется разрешение.

   - profile

   > [!NOTE]
   > Разрешение `User.Read` может быть уже указано по умолчанию. Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с разрешением, которое не требуется вашей надстройке.

1. Нажмите кнопку **Добавить разрешения** в нижней части панели.

1. На этой же странице выберите согласие администратора **гранта \<tenant-name\>** для кнопки, а затем выберите **Да** для подтверждения, которое появится.

## <a name="create-the-office-add-in"></a>Создание надстройки Office

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Начните Visual Studio 2019 г. и выберите **создание нового проекта**.
1. Поиск и выбор **шаблона Excel веб-надстройки**. Затем нажмите кнопку **Далее**. Примечание. SSO работает с любым Office приложением, но для этой статьи будет работать с Excel.
1. Введите имя проекта, например **sso-display-user-info** и выберите **Create**. Другие поля можно оставить по умолчанию.
1. В **диалоговом** окне Выберите диалоговое окно типа надстройки выберите Добавить новые **функции для** Excel и выберите **Finish**.

Проект создан и будет содержать два проекта в решении.

- **sso-display-user-info**: содержит манифест и сведения для боковой загрузки надстройки в Excel.
- **sso-display-user-infoWeb**: проект ASP.NET, на котором размещены веб-страницы надстройки.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Убедитесь, что вы [настроили среду разработки](../overview/set-up-your-dev-environment.md).

1. Чтобы создать проект, введите указанную ниже команду.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

Проект создается в новой папке с именем **sso-display-user-info**.

---

## <a name="configure-the-manifest"></a>Настройка манифеста

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. В **Обозревателе** решений откройте **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. В Visual Studio откройте файл **manifest.xml**.

---

1. В нижней части манифеста находится заключительный `</Resources>` элемент. Вставьте следующий XML чуть ниже элемента `</Resources>` , но перед заключительный `</VersionOverrides>` элемент. Для Office приложений, кроме Outlook, добавьте разметку в конец раздела`<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

   ```xml
   <WebApplicationInfo>
       <Id>[application-id]</Id>
       <Resource>api://localhost:[port]/[application-id]</Resource>
       <Scopes>
           <Scope>openid</Scope>
           <Scope>user.read</Scope>
           <Scope>profile</Scope>
       </Scopes>
   </WebApplicationInfo>
   ```

1. Замените `[port]` правильный номер порта для проекта. Если надстройка создана с помощью yo office, номер порта обычно составляет 3000 и находится в файле package.json. Если надстройка создана с Visual Studio 2019 г., порт находится в **свойстве URL-адреса SSL** веб-проекта.
1. Замените `[application-id]` оба задатки фактическим ИД приложения из регистрации приложения.
1. Сохраните файл.

Вставленный XML содержит следующие элементы и сведения.

- **WebApplicationInfo** — родительский элемент для указанных ниже элементов.
- **Id** - Идентификатор клиента надстройки. Это идентификатор приложения, который вы получаете в процессе регистрации надстройки. См. [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](register-sso-add-in-aad-v2.md).
- **Resource** — URL-адрес надстройки; Это тот же URI (включая протокол `api:`), который вы использовали при регистрации надстройки и в AAD. Доменная часть данного URI должна соответствовать домену, в том числе поддомену, используемом в URL-адресах в части`<Resources>` манифеста настройки. URI должен заканчиваться идентификатором клиента в `<Id>`.
- **Scopes** — родительский элемент одного или нескольких элементов **Scope**;
- **Scope** — указывает разрешение, необходимое надстройке для работы с AAD. Разрешения `profile` и `openID` требуются во всех случаях и они могут быть единственными необходимыми разрешениями, если ваша надстройка не имеет доступа к Microsoft Graph. В противном случае вам также могут потребоваться элементы типа **Область** для необходимым разрешений Microsoft Graph; например, `User.Read`, `Mail.Read`. Библиотеки, которые вы используете в коде, чтобы получить доступ к Microsoft Graph, могут потребовать дополнительные разрешения. Например, библиотека проверки подлинности Microsoft (MSAL) для .NET требует разрешения `offline_access`. Дополнительные сведения см. в статье [Авторизация в Microsoft Graph для надстройки Office](authorize-to-microsoft-graph.md).

## <a name="add-the-jwt-decode-package"></a>Добавление пакета jwt-decode

Вы можете вызвать `getAccessToken` API, чтобы получить маркер ID из Office. Сначала позволяет добавить пакет jwt-decode, чтобы упростить расшифровку и просмотр маркера ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Откройте Visual Studio решение.
1. В меню выберите **Инструменты > NuGet диспетчер пакетов > диспетчер пакетов Консоли**.
1. Введите следующую команду **в консоли диспетчер пакетов консоли**.

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Из окна терминала или консоли перейдите в корневую папку для проекта надстройки.
1. Введите следующую команду

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>Добавление пользовательского интерфейса в области задач

Нам необходимо изменить области задач, чтобы она отображала сведения пользователей, которые мы получаем из маркера ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Откройте файл Home.html.
1. Добавьте следующий тег скрипта в `<head>` раздел страницы. Это будет пакет jwt-decode, который мы добавили ранее.

   ```html
   <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
   ```

1. Замените `<body>` раздел следующим HTML.

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Откройте **файл src/taskpane/taskpane.html** .
1. Замените `<body>` раздел следующим HTML.

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

---

## <a name="call-the-getaccesstoken-api"></a>Вызов API getAccessToken

Заключительный шаг — получить маркер ID по вызову `getAccessToken`.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Откройте файл **Home.js** .
1. Замените все содержимое файла указанным ниже кодом.

   ```javascript
   (function () {
     "use strict";

     // The initialize function must be run each time a new page is loaded.
     Office.initialize = function (reason) {
       $(document).ready(function () {
         $("#getIDToken").click(getIDToken);
       });
     };

     async function getIDToken() {
       try {
         let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
           allowSignInPrompt: true,
         });
         let userToken = jwt_decode(userTokenEncoded);
         document.getElementById("userInfo").innerHTML =
           "name: " +
           userToken.name +
           "<br>email: " +
           userToken.preferred_username +
           "<br>id: " +
           userToken.oid;
         console.log(userToken);
       } catch (error) {
         document.getElementById("userInfo").innerHTML =
           "An error occurred. <br>Name: " +
           error.name +
           "<br>Code: " +
           error.code +
           "<br>Message: " +
           error.message;
         console.log(error);
       }
     }
   })();
   ```

1. Сохраните файл.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Откройте **файл src/taskpane/taskpane.js** .
1. Замените все содержимое файла указанным ниже кодом.

   ```javascript
   import jwt_decode from "jwt-decode";

   Office.onReady((info) => {
     if (info.host === Office.HostType.Excel) {
       document.getElementById("getIDToken").onclick = getIDToken;
     }
   });

   async function getIDToken() {
     try {
       let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
         allowSignInPrompt: true,
       });
       let userToken = jwt_decode(userTokenEncoded);
       document.getElementById("userInfo").innerHTML =
         "name: " +
         userToken.name +
         "<br>email: " +
         userToken.preferred_username +
         "<br>id: " +
         userToken.oid;
       console.log(userToken);
     } catch (error) {
       document.getElementById("userInfo").innerHTML =
         "An error occurred. <br>Name: " +
         error.name +
         "<br>Code: " +
         error.code +
         "<br>Message: " +
         error.message;
       console.log(error);
     }
   }
   ```

1. Сохраните файл.

---

## <a name="run-the-add-in"></a>Запуск надстройки

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Выберите **отладку > начать отладку** или нажмите **кнопку F5**.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Запуск `npm start` из командной строки.

---

1. Когда Excel начинается, Office с той же учетной записью клиента, что и для создания регистрации приложения.
1. На **ленте Главная** выберите **Показать задачу** , чтобы открыть надстройку.
1. В области задач надстройки выберите маркер **Get ID**.

Надстройка отображает имя, электронную почту и ID учетной записи, с помощью которого вы подписались.

> [!NOTE]
> Если вы столкнулись с ошибками, просмотрите этапы регистрации в этой статье для регистрации приложения. Отсутствие подробной информации при настройке регистрации приложения является распространенной причиной проблем, с помощью SSO. Если вы по-прежнему не можете успешно запустить надстройку, см. сообщение об ошибке устранения неполадок для одного входного знака [(SSO)](troubleshoot-sso-in-office-add-ins.md).

## <a name="see-also"></a>См. также

[Использование утверждений для надежной идентификации пользователя (Subject and Object ID)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)
