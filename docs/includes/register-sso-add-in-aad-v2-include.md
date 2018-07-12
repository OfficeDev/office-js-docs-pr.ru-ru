

1. <span data-ttu-id="7d460-101">Перейдите к [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="7d460-101">Navigate to [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)</span></span>

1. <span data-ttu-id="7d460-p101">Войдите в клиент Office 365, используя учетные данные администратора. Пример: MyName@contoso.onmicrosoft.com</span><span class="sxs-lookup"><span data-stu-id="7d460-p101">Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com</span></span>

1. <span data-ttu-id="7d460-104">Нажмите **Добавить приложение**.</span><span class="sxs-lookup"><span data-stu-id="7d460-104">Click **Add an app**.</span></span>

1. <span data-ttu-id="7d460-105">При появлении запроса введите **$ADD-IN-NAME$** в качестве имени приложения, а затем нажмите **Создать приложение**.</span><span class="sxs-lookup"><span data-stu-id="7d460-105">When prompted, use “Office-Add-in-ASPNET-SSO” as the app name, and then press Create application.</span></span>

1. <span data-ttu-id="7d460-p102">Когда откроется страница настройки, скопируйте и сохраните **идентификатор приложения**. Он понадобится вам позже.</span><span class="sxs-lookup"><span data-stu-id="7d460-p102">When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.</span></span>

    > [!NOTE]
    > <span data-ttu-id="7d460-p103">Этот идентификатор представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel и т. д.), пытаются получить авторизованный доступ к вашему приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="7d460-p103">This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="7d460-p104">В разделе **Секреты приложения** нажмите **Создать новый пароль**. Откроется всплывающее диалоговое окно с новым паролем (его также называют секретом приложения). *Сразу скопируйте пароль и сохраните его вместе с идентификатором приложения.* Он понадобится вам позже. Затем закройте диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="7d460-p104">In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.</span></span>

1. <span data-ttu-id="7d460-115">В разделе **Платформы** нажмите **Добавление платформы**.</span><span class="sxs-lookup"><span data-stu-id="7d460-115">In the **Platforms** section, click **Add Platform**.</span></span>

1. <span data-ttu-id="7d460-116">В открывшемся диалоговом окне выберите **Веб-API**.</span><span class="sxs-lookup"><span data-stu-id="7d460-116">In the dialog that opens, select **Web API**.</span></span>

1. <span data-ttu-id="7d460-117">**URI идентификатора приложения** сгенерирован в форме “api://$App ID GUID$”.</span><span class="sxs-lookup"><span data-stu-id="7d460-117">An **Application ID URI** has been generated of the form “api://$App ID GUID$”.</span></span> <span data-ttu-id="7d460-118">Вставьте **$FQDN-БЕЗ-ПРОТОКОЛА$** (с косой чертой "/", на конце) между двойной косой чертой и GUID.</span><span class="sxs-lookup"><span data-stu-id="7d460-118">Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="7d460-119">Весь идентификатор должен иметь форму `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`, например: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="7d460-119">The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="7d460-120">Если возникает ошибка с сообщением о том, что домен уже занят, но при этом вы являетесь его владельцем, следуйте процедуре в статье [Краткое руководство. Добавление личного домена в Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain), чтобы зарегистрировать его, а затем повторите этот шаг.</span><span class="sxs-lookup"><span data-stu-id="7d460-120">If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) to register it, and then repeat this step.</span></span>

    > [!NOTE]
    > <span data-ttu-id="7d460-121">Доменная часть имени в поле**Область**, указанная под**URI идентификатора приложения**, автоматически изменится соответствующим образом с добавлением `/access_as_user`в конце, например:`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="7d460-121">The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match.</span></span>

1. <span data-ttu-id="7d460-122">В разделе **Предварительно авторизованные приложения** укажите приложения, которые необходимо авторизовать для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="7d460-122">In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="7d460-123">Необходимо обеспечить предварительную авторизацию для всех указанных ниже идентификаторов.</span><span class="sxs-lookup"><span data-stu-id="7d460-123">Each of the following IDs needs to be pre-authorized.</span></span> <span data-ttu-id="7d460-124">После ввода каждого из них будет появляться новое пустое текстовое поле.</span><span class="sxs-lookup"><span data-stu-id="7d460-124">Each time you enter one, a new empty textbox appears.</span></span> <span data-ttu-id="7d460-125">(Введите только GUID.)</span><span class="sxs-lookup"><span data-stu-id="7d460-125">(Enter only the GUID.)</span></span>
    * <span data-ttu-id="7d460-126">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="7d460-126">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="7d460-127">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online).</span><span class="sxs-lookup"><span data-stu-id="7d460-127">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span></span>
    * <span data-ttu-id="7d460-128">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online).</span><span class="sxs-lookup"><span data-stu-id="7d460-128">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span></span>

1. <span data-ttu-id="7d460-129">Откройте раскрывающийся список **Область** рядом с каждым полем **Идентификатор приложения** и установите флажок `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="7d460-129">Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span></span>

1. <span data-ttu-id="7d460-130">В верхней части раздела **Платформы** снова нажмите кнопку **Добавление платформы** и выберите **Веб**.</span><span class="sxs-lookup"><span data-stu-id="7d460-130">Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.</span></span>

1. <span data-ttu-id="7d460-131">В новом подразделе **Веб** раздела **Платформы** введите следующий **URL-адрес перенаправления**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span><span class="sxs-lookup"><span data-stu-id="7d460-131">In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span></span>

1. <span data-ttu-id="7d460-p107">Прокрутите вниз до подраздела **Делегированные разрешения** в разделе **Разрешения Microsoft Graph**. Нажмите кнопку **Добавить**, чтобы открыть диалоговое окно **Выбор разрешения**.</span><span class="sxs-lookup"><span data-stu-id="7d460-p107">Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.</span></span>

1. <span data-ttu-id="7d460-134">В диалоговом окне установите флажки для `profile` и других разрешений AAD и Microsoft Graph, необходимых вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="7d460-134">In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs.</span></span> <span data-ttu-id="7d460-135">Примеры:</span><span class="sxs-lookup"><span data-stu-id="7d460-135">The following are examples:</span></span>

    * <span data-ttu-id="7d460-136">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="7d460-136">Files.Read.All</span></span>
    * <span data-ttu-id="7d460-137">offline_access</span><span class="sxs-lookup"><span data-stu-id="7d460-137">offline_access</span></span>
    * <span data-ttu-id="7d460-138">openid</span><span class="sxs-lookup"><span data-stu-id="7d460-138">openid</span></span>
    * <span data-ttu-id="7d460-139">профиль</span><span class="sxs-lookup"><span data-stu-id="7d460-139">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="7d460-140">Разрешение `User.Read` может быть уже указано по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="7d460-140">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="7d460-141">Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с этим разрешением, если оно не требуется вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="7d460-141">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.</span></span>

1. <span data-ttu-id="7d460-142">Нажмите кнопку **ОК** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="7d460-142">At the bottom of the dialog, click **OK**.</span></span>

1. <span data-ttu-id="7d460-143">Нажмите кнопку **Сохранить** в нижней части страницы регистрации.</span><span class="sxs-lookup"><span data-stu-id="7d460-143">At the bottom of the registration page, click **Save**.</span></span>
