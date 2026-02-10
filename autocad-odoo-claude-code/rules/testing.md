# 測試規範

> **優先級**: HIGH
> **適用**: 所有代理

---

## 1. TDD 工作流

嚴格遵循 **Red → Green → Refactor** 循環：

1. **Red**: 撰寫一個描述期望行為的失敗測試
2. **Green**: 實作最少的程式碼使測試通過
3. **Refactor**: 改善程式碼結構，保持測試通過
4. **Build**: 驗證整體建置通過
5. **Commit**: 行為變更和結構變更分開提交

## 2. 測試組織

```
tests/OdooAutoCAD.Integration.Tests/
├── ViewModels/              # ViewModel 測試
│   ├── MainViewModelTests.cs
│   └── DashboardViewModelTests.cs
├── Services/                # 服務測試
│   ├── NavigationServiceTests.cs
│   └── AutoCADServiceTests.cs
├── Converters/              # 轉換器測試
│   └── EqualityConverterTests.cs
├── Integration/             # 整合測試
│   ├── ShutdownCleanupTests.cs
│   └── NavigationFlowTests.cs
└── Helpers/                 # 測試輔助工具
    └── TestServiceProvider.cs
```

## 3. 測試命名規範

```
{被測方法}_{測試場景}_{期望結果}
```

範例：
- `NavigateTo_ValidPage_ReturnsPageInstance`
- `NavigateTo_NullPageName_ThrowsArgumentException`
- `Connect_AutoCADNotRunning_ReturnsFalse`
- `Shutdown_WithTimeout_CompletesWithinFiveSeconds`

## 4. 測試模板

### 單元測試
```csharp
public class MainViewModelTests
{
    private readonly Mock<INavigationService> _mockNav;
    private readonly MainViewModel _sut; // System Under Test

    public MainViewModelTests()
    {
        _mockNav = new Mock<INavigationService>();
        _sut = new MainViewModel(_mockNav.Object);
    }

    [Fact]
    public void NavigateCommand_ValidPage_CallsNavigationService()
    {
        // Arrange
        var page = "Dashboard";

        // Act
        _sut.NavigateCommand.Execute(page);

        // Assert
        _mockNav.Verify(n => n.NavigateTo(It.IsAny<string>()), Times.Once);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void NavigateCommand_InvalidPage_DoesNotNavigate(string? page)
    {
        // Act
        _sut.NavigateCommand.Execute(page);

        // Assert
        _mockNav.Verify(n => n.NavigateTo(It.IsAny<string>()), Times.Never);
    }
}
```

### 整合測試
```csharp
public class ShutdownCleanupTests
{
    [Fact]
    public async Task Shutdown_AllSteps_CompletesWithinTimeout()
    {
        // Arrange
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

        // Act & Assert
        await Assert.CompletesWithinAsync(
            TimeSpan.FromSeconds(5),
            () => _app.ShutdownAsync(cts.Token));
    }
}
```

## 5. 覆蓋率目標

| 層級 | 目標 | 說明 |
|------|------|------|
| ViewModel | 90%+ | 核心業務邏輯 |
| Services | 85%+ | 服務層邏輯 |
| Converters | 95%+ | 值轉換器（純函式） |
| Integration | 80%+ | 端對端流程 |

## 6. Mock 規則

- 外部依賴（AutoCAD COM、Odoo API）**必須** Mock
- 使用介面（`INavigationService`）方便 Mock 注入
- Mock 行為只設定測試需要的部分

## 7. 測試指令

```bash
# 執行所有測試
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --verbosity normal

# 特定測試
"C:\Program Files\dotnet\dotnet.exe" test --filter "FullyQualifiedName~ShutdownCleanup"

# 覆蓋率報告（需要 coverlet）
"C:\Program Files\dotnet\dotnet.exe" test --collect:"XPlat Code Coverage"
```

## 測試檢查清單

- [ ] 新功能有對應測試
- [ ] Bug 修復有回歸測試
- [ ] 測試可獨立執行（無順序依賴）
- [ ] 外部依賴已 Mock
- [ ] 測試命名遵循規範
- [ ] 所有測試通過後才能提交
