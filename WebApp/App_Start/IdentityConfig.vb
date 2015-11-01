Imports System.Security.Claims
Imports System.Threading.Tasks
Imports Microsoft.AspNet.Identity
Imports Microsoft.AspNet.Identity.EntityFramework
Imports Microsoft.AspNet.Identity.Owin
Imports Microsoft.Owin
Imports Microsoft.Owin.Security

Public Class EmailService
    Implements IIdentityMessageService
    Public Function SendAsync(message As IdentityMessage) As Task Implements IIdentityMessageService.SendAsync
        ' 電子メールを送信するには、電子メール サービスをここにプラグインします。
        Return Task.FromResult(0)
    End Function
End Class

Public Class SmsService
    Implements IIdentityMessageService
    Public Function SendAsync(message As IdentityMessage) As Task Implements IIdentityMessageService.SendAsync
        ' テキスト メッセージを送信するには、SMS サービスをここにプラグインします。
        Return Task.FromResult(0)
    End Function
End Class

' このアプリケーションで使用されるアプリケーション ユーザー マネージャーを設定します。UserManager は ASP.NET Identity の中で定義されており、このアプリケーションで使用されます。
Public Class ApplicationUserManager
    Inherits UserManager(Of ApplicationUser)
    Public Sub New(store As IUserStore(Of ApplicationUser))
        MyBase.New(store)
    End Sub

    Public Shared Function Create(options As IdentityFactoryOptions(Of ApplicationUserManager), context As IOwinContext) As ApplicationUserManager
        Dim manager = New ApplicationUserManager(New UserStore(Of ApplicationUser)(context.[Get](Of ApplicationDbContext)()))
        ' ユーザー名の検証ロジックを設定します
        manager.UserValidator = New UserValidator(Of ApplicationUser)(manager) With {
          .AllowOnlyAlphanumericUserNames = False,
          .RequireUniqueEmail = True
        }

        ' パスワードの検証ロジックを設定します
        manager.PasswordValidator = New PasswordValidator() With {
          .RequiredLength = 6,
          .RequireNonLetterOrDigit = True,
          .RequireDigit = True,
          .RequireLowercase = True,
          .RequireUppercase = True
        }
        ' 2 要素認証プロバイダーを登録します。このアプリケーションでは、Phone and Emails をユーザー検証用コード受け取りのステップとして使用します。
        ' 独自のプロバイダーを作成して、ここにプラグインすることができます。
        manager.RegisterTwoFactorProvider("コードを電話で伝える", New PhoneNumberTokenProvider(Of ApplicationUser)() With {
          .MessageFormat = "セキュリティ コードは {0} です"
        })
        manager.RegisterTwoFactorProvider("コードをメールで送信する", New EmailTokenProvider(Of ApplicationUser)() With {
          .Subject = "セキュリティ コード",
          .BodyFormat = "セキュリティ コードは {0} です"
        })
        
        ' ユーザー ロックアウトの既定値を設定
        manager.UserLockoutEnabledByDefault = True
        manager.DefaultAccountLockoutTimeSpan = TimeSpan.FromMinutes(5)
        manager.MaxFailedAccessAttemptsBeforeLockout = 5
        
        manager.EmailService = New EmailService()
        manager.SmsService = New SmsService()
        Dim dataProtectionProvider = options.DataProtectionProvider
        If dataProtectionProvider IsNot Nothing Then
            manager.UserTokenProvider = New DataProtectorTokenProvider(Of ApplicationUser)(dataProtectionProvider.Create("ASP.NET Identity"))
        End If
        Return manager
    End Function
End Class

Public Class ApplicationSignInManager
    Inherits SignInManager(Of ApplicationUser, String)
    Public Sub New(userManager As ApplicationUserManager, authenticationManager As IAuthenticationManager)
        MyBase.New(userManager, authenticationManager)
    End Sub

    Public Overrides Function CreateUserIdentityAsync(user As ApplicationUser) As Task(Of ClaimsIdentity)
        Return user.GenerateUserIdentityAsync(DirectCast(UserManager, ApplicationUserManager))
    End Function

    Public Shared Function Create(options As IdentityFactoryOptions(Of ApplicationSignInManager), context As IOwinContext) As ApplicationSignInManager
        Return New ApplicationSignInManager(context.GetUserManager(Of ApplicationUserManager)(), context.Authentication)
    End Function
End Class
