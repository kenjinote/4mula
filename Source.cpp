#define NOMINMAX

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")

#include <windows.h>
#include <wrl.h>
#include <d2d1_3.h>
#include <dwrite_3.h>
#include <string>
#include <vector>
#include <memory>
#include <algorithm>

using namespace Microsoft::WRL;

// ============================================================================
// 1. 描画インターフェース (APIの抽象化)
// ============================================================================
class IMathRendererContext {
public:
    virtual ~IMathRendererContext() = default;
    virtual void DrawGlyph(const std::wstring& text, float x, float y, float fontSize, bool isItalic) = 0;
    virtual void DrawLine(float x1, float y1, float x2, float y2, float thickness) = 0;
    // TeXの寸法（幅、ベースラインからの高さ、ベースラインからの深さ）を測定
    virtual void MeasureGlyph(const std::wstring& text, float fontSize, bool isItalic, float& outWidth, float& outHeight, float& outDepth) = 0;
};

// ============================================================================
// 2. Boxモデル (レイアウトの最小単位)
// ============================================================================
class Box {
public:
    float width = 0.0f;
    float height = 0.0f; // ベースラインより上の高さ
    float depth = 0.0f;  // ベースラインより下の深さ
    float shift = 0.0f;  // 微調整用Yオフセット (上付き文字などに使用)

    virtual ~Box() = default;
    virtual void Draw(IMathRendererContext* context, float x, float y) const = 0;
};

class CharBox : public Box {
    std::wstring character;
    float fontSize;
    bool isItalic;
public:
    CharBox(const std::wstring& c, float size, bool italic, IMathRendererContext* ctx)
        : character(c), fontSize(size), isItalic(italic) {
        ctx->MeasureGlyph(character, fontSize, isItalic, width, height, depth);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        context->DrawGlyph(character, x, y + shift, fontSize, isItalic);
    }
};

class SpaceBox : public Box {
public:
    SpaceBox(float w) { width = w; height = 0; depth = 0; }
    void Draw(IMathRendererContext*, float, float) const override {}
};

class HorizontalBox : public Box {
    std::vector<std::shared_ptr<Box>> children;
public:
    void Add(std::shared_ptr<Box> box) {
        children.push_back(box);
        width += box->width;
        height = std::max(height, box->height - box->shift);
        depth = std::max(depth, box->depth + box->shift);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentX = x;
        for (const auto& child : children) {
            child->Draw(context, currentX, y);
            currentX += child->width;
        }
    }
};

// 分数をレイアウトするBox
class FractionBox : public Box {
    std::shared_ptr<Box> numerator;
    std::shared_ptr<Box> denominator;
    float paddingX = 4.0f;        // 分数線の左右のはみ出し
    float mathAxisOffset = 12.0f; // ★修正: イコールの中心(数式軸)の高さ。フォントサイズ36の場合は約12px
    float gap = 5.0f;             // ★追加: 分数線と分子・分母の隙間
public:
    FractionBox(std::shared_ptr<Box> num, std::shared_ptr<Box> den) : numerator(num), denominator(den) {
        width = std::max(num->width, den->width) + paddingX * 2.0f;
        // 高さ・深さも数式軸を基準に正確に再計算
        height = mathAxisOffset + gap + num->depth + num->height;
        depth = (den->height + gap + den->depth) - mathAxisOffset;
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float mathAxis = y - mathAxisOffset; // ★修正: 数式軸を正確な位置に

        // 分子を中央揃えで描画
        float numX = x + (width - numerator->width) / 2.0f;
        float numBaseline = mathAxis - gap - numerator->depth; // ★修正: ベースラインを正確に計算
        numerator->Draw(context, numX, numBaseline);

        // 分母を中央揃えで描画
        float denX = x + (width - denominator->width) / 2.0f;
        float denBaseline = mathAxis + gap + denominator->height; // ★修正: ベースラインを正確に計算
        denominator->Draw(context, denX, denBaseline);

        // 分数線を描画
        context->DrawLine(x, mathAxis, x + width, mathAxis, 1.5f);
    }
};

// 文字列をまとめて描画するBox (sin, cos, dx などに便利)
class StringBox : public Box {
    std::wstring text;
    float fontSize;
    bool isItalic;
public:
    StringBox(const std::wstring& t, float size, bool italic, IMathRendererContext* ctx)
        : text(t), fontSize(size), isItalic(italic) {
        ctx->MeasureGlyph(text, fontSize, isItalic, width, height, depth);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        context->DrawGlyph(text, x, y + shift, fontSize, isItalic);
    }
};

// 上付き・下付き文字を管理する汎用Box (指数関数やインテグラルに使用)
class ScriptBox : public Box {
    std::shared_ptr<Box> baseBox;
    std::shared_ptr<Box> superscript;
    std::shared_ptr<Box> subscript;
public:
    ScriptBox(std::shared_ptr<Box> base, std::shared_ptr<Box> sup, std::shared_ptr<Box> sub)
        : baseBox(base), superscript(sup), subscript(sub) {

        width = base->width;
        float scriptWidth = 0.0f;

        if (superscript) scriptWidth = std::max(scriptWidth, superscript->width);
        if (subscript)   scriptWidth = std::max(scriptWidth, subscript->width);

        width += scriptWidth + 2.0f; // 添字との間にわずかなマージン

        height = base->height;
        depth = base->depth;

        // 上付き文字のYオフセット計算 (ベース文字の高さの約60%上に配置)
        if (superscript) {
            superscript->shift = -base->height * 0.6f;
            height = std::max(height, superscript->height - superscript->shift);
        }
        // 下付き文字のYオフセット計算 (ベース文字の深さに合わせて下に配置)
        if (subscript) {
            subscript->shift = base->depth + (subscript->height * 0.4f);
            depth = std::max(depth, subscript->depth + subscript->shift);
        }
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        baseBox->Draw(context, x, y);
        float scriptX = x + baseBox->width + 2.0f; // ベース文字の右隣に配置
        if (superscript) superscript->Draw(context, scriptX, y);
        if (subscript)   subscript->Draw(context, scriptX, y);
    }
};

// ============================================================================
// 3. Direct2D / DirectWrite 実装コンテキスト
// ============================================================================
class Direct2DContext : public IMathRendererContext {
    ComPtr<ID2D1RenderTarget> m_pRT;
    ComPtr<IDWriteFactory3> m_pDWriteFactory;
    ComPtr<IDWriteTextFormat> m_pFormatNormal;
    ComPtr<IDWriteTextFormat> m_pFormatItalic;
    ComPtr<ID2D1SolidColorBrush> m_pBrush;

public:
    Direct2DContext(ID2D1RenderTarget* pRT, IDWriteFactory3* pDWriteFactory)
        : m_pRT(pRT), m_pDWriteFactory(pDWriteFactory) {
        m_pRT->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), &m_pBrush);

        // Cambria Math (立体)
        m_pDWriteFactory->CreateTextFormat(L"Cambria Math", nullptr, DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 32.0f, L"en-us", &m_pFormatNormal);
        // Cambria Math (斜体)
        m_pDWriteFactory->CreateTextFormat(L"Cambria Math", nullptr, DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_ITALIC, DWRITE_FONT_STRETCH_NORMAL, 32.0f, L"en-us", &m_pFormatItalic);
    }

    void MeasureGlyph(const std::wstring& text, float fontSize, bool isItalic, float& outWidth, float& outHeight, float& outDepth) override {
        ComPtr<IDWriteTextLayout> pLayout;
        IDWriteTextFormat* format = isItalic ? m_pFormatItalic.Get() : m_pFormatNormal.Get();

        m_pDWriteFactory->CreateTextLayout(text.c_str(), static_cast<UINT32>(text.length()), format, 10000.0f, 10000.0f, &pLayout);
        DWRITE_TEXT_RANGE range = { 0, static_cast<UINT32>(text.length()) };
        pLayout->SetFontSize(fontSize, range);

        // 全体の幅と高さを取得
        DWRITE_TEXT_METRICS metrics;
        pLayout->GetMetrics(&metrics);
        outWidth = metrics.widthIncludingTrailingWhitespace;

        // ベースラインを取得して、Height(上)とDepth(下)に分割する
        DWRITE_LINE_METRICS lineMetrics;
        UINT32 actualLineCount;
        pLayout->GetLineMetrics(&lineMetrics, 1, &actualLineCount);

        outHeight = lineMetrics.baseline;
        outDepth = lineMetrics.height - lineMetrics.baseline;
    }

    void DrawGlyph(const std::wstring& text, float x, float y, float fontSize, bool isItalic) override {
        ComPtr<IDWriteTextLayout> pLayout;
        IDWriteTextFormat* format = isItalic ? m_pFormatItalic.Get() : m_pFormatNormal.Get();

        m_pDWriteFactory->CreateTextLayout(text.c_str(), static_cast<UINT32>(text.length()), format, 10000.0f, 10000.0f, &pLayout);
        DWRITE_TEXT_RANGE range = { 0, static_cast<UINT32>(text.length()) };
        pLayout->SetFontSize(fontSize, range);

        // y座標はベースライン位置なので、描画開始位置(左上)を計算して渡す
        DWRITE_LINE_METRICS lineMetrics;
        UINT32 actualLineCount;
        pLayout->GetLineMetrics(&lineMetrics, 1, &actualLineCount);

        m_pRT->DrawTextLayout(D2D1::Point2F(x, y - lineMetrics.baseline), pLayout.Get(), m_pBrush.Get());
    }

    void DrawLine(float x1, float y1, float x2, float y2, float thickness) override {
        m_pRT->DrawLine(D2D1::Point2F(x1, y1), D2D1::Point2F(x2, y2), m_pBrush.Get(), thickness);
    }
};

// ============================================================================
// 4. Windows アプリケーション エントリーポイント
// ============================================================================
ComPtr<ID2D1Factory7> g_pD2DFactory;
ComPtr<IDWriteFactory3> g_pDWriteFactory;
ComPtr<ID2D1HwndRenderTarget> g_pRenderTarget;
std::shared_ptr<Box> g_pLayoutTree;

// 構文木(AST)を手動で構築するデモ (修正版)
std::shared_ptr<Box> BuildDemoFormula(IMathRendererContext* ctx) {
    auto root = std::make_shared<HorizontalBox>();
    float normalSize = 30.0f;
    float scriptSize = 18.0f; // 添字は少し小さく
    float integralSize = 50.0f; // インテグラルは大きく

    // 1. 微分記号: d/dx
    auto d_num = std::make_shared<CharBox>(L"d", normalSize, false, ctx); // 微分のdは立体が好まれる
    auto d_den = std::make_shared<StringBox>(L"dx", normalSize, false, ctx);
    root->Add(std::make_shared<FractionBox>(d_num, d_den));
    root->Add(std::make_shared<SpaceBox>(10.0f));

    // 2. 積分記号: \int_0^x
    auto intSym = std::make_shared<CharBox>(L"\x222B", integralSize, false, ctx); // U+222B がインテグラル
    auto intSub = std::make_shared<CharBox>(L"0", scriptSize, false, ctx);
    auto intSup = std::make_shared<CharBox>(L"x", scriptSize, true, ctx);
    root->Add(std::make_shared<ScriptBox>(intSym, intSup, intSub));
    root->Add(std::make_shared<SpaceBox>(8.0f));

    // 3. 三角関数: sin(t)
    root->Add(std::make_shared<StringBox>(L"sin", normalSize, false, ctx)); // sinは立体
    root->Add(std::make_shared<CharBox>(L"(", normalSize, false, ctx));
    root->Add(std::make_shared<CharBox>(L"t", normalSize, true, ctx));
    root->Add(std::make_shared<CharBox>(L")", normalSize, false, ctx));
    root->Add(std::make_shared<SpaceBox>(4.0f));

    // 4. 指数関数: e^{-t}
    auto e_base = std::make_shared<CharBox>(L"e", normalSize, true, ctx); // ネイピア数eは斜体
    auto e_sup = std::make_shared<HorizontalBox>();
    e_sup->Add(std::make_shared<CharBox>(L"\x2212", scriptSize, false, ctx)); // 正しいマイナス記号
    e_sup->Add(std::make_shared<CharBox>(L"t", scriptSize, true, ctx));
    root->Add(std::make_shared<ScriptBox>(e_base, e_sup, nullptr)); // 下付きはnullptr
    root->Add(std::make_shared<SpaceBox>(8.0f));

    // 5. dt
    root->Add(std::make_shared<StringBox>(L"dt", normalSize, false, ctx));
    root->Add(std::make_shared<SpaceBox>(14.0f));

    // 6. イコール
    root->Add(std::make_shared<CharBox>(L"=", normalSize, false, ctx));
    root->Add(std::make_shared<SpaceBox>(14.0f));

    // 7. 右辺: sin(x) e^{-x}
    root->Add(std::make_shared<StringBox>(L"sin", normalSize, false, ctx));
    root->Add(std::make_shared<CharBox>(L"(", normalSize, false, ctx));
    root->Add(std::make_shared<CharBox>(L"x", normalSize, true, ctx));
    root->Add(std::make_shared<CharBox>(L")", normalSize, false, ctx));
    root->Add(std::make_shared<SpaceBox>(4.0f));

    auto ex_base = std::make_shared<CharBox>(L"e", normalSize, true, ctx);
    auto ex_sup = std::make_shared<HorizontalBox>();
    ex_sup->Add(std::make_shared<CharBox>(L"\x2212", scriptSize, false, ctx));
    ex_sup->Add(std::make_shared<CharBox>(L"x", scriptSize, true, ctx));
    root->Add(std::make_shared<ScriptBox>(ex_base, ex_sup, nullptr));

    return root;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        BeginPaint(hWnd, &ps);

        if (!g_pRenderTarget) {
            RECT rc;
            GetClientRect(hWnd, &rc);
            D2D1_SIZE_U size = D2D1::SizeU(rc.right - rc.left, rc.bottom - rc.top);
            g_pD2DFactory->CreateHwndRenderTarget(D2D1::RenderTargetProperties(), D2D1::HwndRenderTargetProperties(hWnd, size), &g_pRenderTarget);

            // 初回のみレイアウトツリーを構築
            Direct2DContext ctx(g_pRenderTarget.Get(), g_pDWriteFactory.Get());
            g_pLayoutTree = BuildDemoFormula(&ctx);
        }

        g_pRenderTarget->BeginDraw();
        g_pRenderTarget->Clear(D2D1::ColorF(D2D1::ColorF::White));

        if (g_pLayoutTree) {
            Direct2DContext ctx(g_pRenderTarget.Get(), g_pDWriteFactory.Get());
            // 余白を取って描画 (x=100, y=200のベースライン位置)
            g_pLayoutTree->Draw(&ctx, 100.0f, 200.0f);
        }

        g_pRenderTarget->EndDraw();
        EndPaint(hWnd, &ps);
    }
    break;
    case WM_SIZE:
        if (g_pRenderTarget) {
            RECT rc;
            GetClientRect(hWnd, &rc);
            g_pRenderTarget->Resize(D2D1::SizeU(rc.right - rc.left, rc.bottom - rc.top));
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow) {
    D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, g_pD2DFactory.GetAddressOf());
    DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory3), reinterpret_cast<IUnknown**>(g_pDWriteFactory.GetAddressOf()));

    WNDCLASSEXW wcex = { sizeof(WNDCLASSEX) };
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.lpszClassName = L"4mulaWindow";
    RegisterClassExW(&wcex);

    HWND hWnd = CreateWindowW(L"4mulaWindow", L"4mula - Box Model Rendering", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, 800, 450, nullptr, nullptr, hInstance, nullptr);
    if (!hWnd) return FALSE;

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}