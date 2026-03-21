#define NOMINMAX
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "dwmapi.lib")
#include <windows.h>
#include <windowsx.h>
#include <wrl.h>
#include <d2d1_3.h>
#include <dwrite_3.h>
#include <dwmapi.h>
#include <string>
#include <vector>
#include <memory>
#include <cmath>
#include <algorithm>
#include <stack>
#include <sstream>
#include <iomanip>
#include <boost/multiprecision/cpp_dec_float.hpp>
#include <boost/math/special_functions/gamma.hpp>
using namespace boost::multiprecision;
typedef boost::multiprecision::number<boost::multiprecision::cpp_dec_float<500>> BigFloat;
using namespace Microsoft::WRL;
class Box;
std::wstring g_resultText = L"";
std::shared_ptr<Box> g_pResultTree = nullptr;
bool g_isDragging = false;
bool g_isDarkMode = false;
class IMathRendererContext {
public:
    virtual ~IMathRendererContext() = default;
    virtual void DrawGlyph(const std::wstring& text, float x, float y, float fontSize, bool isItalic) = 0;
    virtual void DrawLine(float x1, float y1, float x2, float y2, float thickness) = 0;
    virtual void FillCircle(float x, float y, float r) = 0;
    virtual void MeasureGlyph(const std::wstring& text, float fontSize, bool isItalic, float& outWidth, float& outHeight, float& outDepth) = 0;
    virtual void SetTextColor(float r, float g, float b, float a = 1.0f) = 0;
};
struct CaretMetrics {
    float x = -1.0f;
    float y = 0.0f;
    float scale = 1.0f;
    bool isActive = false;
};
std::vector<CaretMetrics> g_caretMetrics;
class Box {
public:
    float width = 0.0f;
    float height = 0.0f;
    float depth = 0.0f;
    float shift = 0.0f;
    int srcStart = -1;
    int srcEnd = -1;
    virtual ~Box() = default;
    virtual void Draw(IMathRendererContext* context, float x, float y) const = 0;
    virtual void SetFontSize(float newSize, IMathRendererContext* ctx) {}
    virtual void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const {}
};
class CharBox : public Box {
public:
    std::wstring character;
    float fontSize;
    bool isItalic;
    CharBox(const std::wstring& c, float size, bool italic, IMathRendererContext* ctx)
        : character(c), fontSize(size), isItalic(italic) {
        ctx->MeasureGlyph(character, fontSize, isItalic, width, height, depth);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        context->DrawGlyph(character, x, y + shift, fontSize, isItalic);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        fontSize = newSize;
        ctx->MeasureGlyph(character, fontSize, isItalic, width, height, depth);
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        if (srcStart >= 0 && srcEnd >= srcStart && srcStart < metrics.size()) {
            if (srcEnd - srcStart == 1) {
                if (metrics[srcStart].x < 0) {
                    metrics[srcStart].x = x; metrics[srcStart].y = y + shift;
                    metrics[srcStart].scale = scale; metrics[srcStart].isActive = isActive;
                }
                if (srcEnd < metrics.size() && metrics[srcEnd].x < 0) {
                    metrics[srcEnd].x = x + width; metrics[srcEnd].y = y + shift;
                    metrics[srcEnd].scale = scale; metrics[srcEnd].isActive = isActive;
                }
            }
            else {
                float cx = x;
                for (int i = 0; i < character.length(); ++i) {
                    if (srcStart + i < metrics.size() && metrics[srcStart + i].x < 0) {
                        metrics[srcStart + i].x = cx; metrics[srcStart + i].y = y + shift;
                        metrics[srcStart + i].scale = scale; metrics[srcStart + i].isActive = isActive;
                    }
                    float cw, ch, cd;
                    ctx->MeasureGlyph(std::wstring(1, character[i]), fontSize, isItalic, cw, ch, cd);
                    cx += cw;
                }
                if (srcEnd < metrics.size() && metrics[srcEnd].x < 0) {
                    metrics[srcEnd].x = x + width; metrics[srcEnd].y = y + shift;
                    metrics[srcEnd].scale = scale; metrics[srcEnd].isActive = isActive;
                }
            }
        }
    }
};
class SpaceBox : public Box {
    float emRatio;
public:
    SpaceBox(float w, float fontSize = 36.0f) {
        emRatio = w / fontSize;
        width = w; height = 0; depth = 0;
    }
    void Draw(IMathRendererContext*, float, float) const override {}
    void SetFontSize(float newSize, IMathRendererContext*) override {
        width = newSize * emRatio;
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        if (srcStart >= 0 && srcStart < metrics.size()) {
            if (metrics[srcStart].x < 0) {
                metrics[srcStart].x = x; metrics[srcStart].y = y + shift;
                metrics[srcStart].scale = scale; metrics[srcStart].isActive = isActive;
            }
            if (srcEnd >= 0 && srcEnd < metrics.size() && metrics[srcEnd].x < 0) {
                metrics[srcEnd].x = x + width; metrics[srcEnd].y = y + shift;
                metrics[srcEnd].scale = scale; metrics[srcEnd].isActive = isActive;
            }
        }
    }
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
        float currentX = x; float currentY = y + shift;
        for (const auto& child : children) {
            child->Draw(context, currentX, currentY); currentX += child->width;
        }
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        width = 0.0f; height = 0.0f; depth = 0.0f;
        for (auto& child : children) {
            child->SetFontSize(newSize, ctx);
            width += child->width;
            height = std::max(height, child->height - child->shift);
            depth = std::max(depth, child->depth + child->shift);
        }
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        if (children.empty() && srcStart >= 0 && srcStart < metrics.size()) {
            if (metrics[srcStart].x < 0) {
                metrics[srcStart].x = x;
                metrics[srcStart].y = y + shift;
                metrics[srcStart].scale = scale;
                metrics[srcStart].isActive = isActive;
            }
            if (srcEnd >= 0 && srcEnd < metrics.size() && metrics[srcEnd].x < 0) {
                metrics[srcEnd].x = x;
                metrics[srcEnd].y = y + shift;
                metrics[srcEnd].scale = scale;
                metrics[srcEnd].isActive = isActive;
            }
        }
        float currentX = x;
        float currentY = y + shift;
        for (const auto& child : children) {
            child->MapCaret(ctx, currentX, currentY, scale, isActive, metrics);
            currentX += child->width;
        }
    }
};
class FractionBox : public Box {
    std::shared_ptr<Box> numerator;
    std::shared_ptr<Box> denominator;
    float paddingX = 4.0f;
    float mathAxisOffset = 12.0f;
    float gap = 5.0f;
    float actualDenHeight = 0.0f;
public:
    FractionBox(std::shared_ptr<Box> num, std::shared_ptr<Box> den) : numerator(num), denominator(den) {
        UpdateMetrics(36.0f);
    }
    void UpdateMetrics(float fontSize) {
        paddingX = fontSize * (4.0f / 36.0f);
        mathAxisOffset = fontSize * (12.0f / 36.0f);
        gap = fontSize * (5.0f / 36.0f);
        float minDenHeight = fontSize * (26.0f / 36.0f);
        actualDenHeight = std::max(denominator->height, minDenHeight);
        width = std::max(numerator->width, denominator->width) + paddingX * 2.0f;
        height = mathAxisOffset + gap + numerator->depth + numerator->height;
        depth = (actualDenHeight + gap + denominator->depth) - mathAxisOffset;
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift;
        float mathAxis = currentY - mathAxisOffset;
        float numX = x + (width - numerator->width) / 2.0f;
        numerator->Draw(context, numX, mathAxis - gap - numerator->depth);
        float denX = x + (width - denominator->width) / 2.0f;
        denominator->Draw(context, denX, mathAxis + gap + actualDenHeight);
        context->DrawLine(x, mathAxis, x + width, mathAxis, 1.5f);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        numerator->SetFontSize(newSize, ctx);
        denominator->SetFontSize(newSize, ctx);
        UpdateMetrics(newSize);
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        float currentY = y + shift;
        float mathAxis = currentY - mathAxisOffset;
        float numX = x + (width - numerator->width) / 2.0f;
        numerator->MapCaret(ctx, numX, mathAxis - gap - numerator->depth, scale, true, metrics);
        float denX = x + (width - denominator->width) / 2.0f;
        denominator->MapCaret(ctx, denX, mathAxis + gap + actualDenHeight, scale, true, metrics);
    }
};
class ScriptBox : public Box {
    std::shared_ptr<Box> baseBox; std::shared_ptr<Box> superscript; std::shared_ptr<Box> subscript;
public:
    ScriptBox(std::shared_ptr<Box> base, std::shared_ptr<Box> sup, std::shared_ptr<Box> sub)
        : baseBox(base), superscript(sup), subscript(sub) {
        RecalculateDimensions();
    }
    void RecalculateDimensions() {
        width = baseBox->width; float scriptWidth = 0.0f;
        if (superscript) scriptWidth = std::max(scriptWidth, superscript->width);
        if (subscript)   scriptWidth = std::max(scriptWidth, subscript->width);
        width += scriptWidth + 2.0f; height = baseBox->height; depth = baseBox->depth;
        if (superscript) {
            superscript->shift = -baseBox->height * 0.6f;
            height = std::max(height, superscript->height - superscript->shift);
        }
        if (subscript) {
            subscript->shift = baseBox->depth + (subscript->height * 0.4f);
            depth = std::max(depth, subscript->depth + subscript->shift);
        }
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift; baseBox->Draw(context, x, currentY);
        float scriptX = x + baseBox->width + 2.0f;
        if (superscript) superscript->Draw(context, scriptX, currentY);
        if (subscript)   subscript->Draw(context, scriptX, currentY);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        baseBox->SetFontSize(newSize, ctx);
        if (superscript) superscript->SetFontSize(newSize * 0.65f, ctx);
        if (subscript) subscript->SetFontSize(newSize * 0.65f, ctx);
        RecalculateDimensions();
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        float currentY = y + shift;
        baseBox->MapCaret(ctx, x, currentY, scale, isActive, metrics);
        float scriptX = x + baseBox->width + 2.0f;
        if (superscript) superscript->MapCaret(ctx, scriptX, currentY, scale * 0.65f, true, metrics);
        if (subscript) subscript->MapCaret(ctx, scriptX, currentY, scale * 0.65f, true, metrics);
    }
};
class StringBox : public Box {
public:
    std::wstring text; float fontSize; bool isItalic;
    StringBox(const std::wstring& t, float size, bool italic, IMathRendererContext* ctx)
        : text(t), fontSize(size), isItalic(italic) {
        ctx->MeasureGlyph(text, fontSize, isItalic, width, height, depth);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override { context->DrawGlyph(text, x, y + shift, fontSize, isItalic); }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        fontSize = newSize; ctx->MeasureGlyph(text, fontSize, isItalic, width, height, depth);
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        if (srcStart >= 0 && srcEnd >= srcStart && srcStart < metrics.size()) {
            if (metrics[srcStart].x < 0) {
                metrics[srcStart].x = x; metrics[srcStart].y = y + shift;
                metrics[srcStart].scale = scale; metrics[srcStart].isActive = isActive;
            }
            if (srcEnd < metrics.size() && metrics[srcEnd].x < 0) {
                metrics[srcEnd].x = x + width; metrics[srcEnd].y = y + shift;
                metrics[srcEnd].scale = scale; metrics[srcEnd].isActive = isActive;
            }
        }
    }
};
class DotBox : public Box {
    std::shared_ptr<Box> innerBox;
public:
    DotBox(std::shared_ptr<Box> inner) : innerBox(inner) {
        width = innerBox->width;
        height = innerBox->height + 10.0f;
        depth = innerBox->depth;
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift;
        innerBox->Draw(context, x, currentY);
        float dotX = x + width / 2.0f;
        float dotY = currentY - innerBox->height - 6.0f;
        context->FillCircle(dotX, dotY, 1.5f);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        innerBox->SetFontSize(newSize, ctx);
        width = innerBox->width;
        height = innerBox->height + (newSize * (10.0f / 36.0f));
        depth = innerBox->depth;
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        innerBox->MapCaret(ctx, x, y + shift, scale, isActive, metrics);
    }
};
class RootBox : public Box {
    std::shared_ptr<Box> innerBox;
    float padLeft = 14.0f; float padTop = 3.0f; float padRight = 2.0f;
    float actualOpeningHeight = 0.0f; float actualOpeningDepth = 0.0f;
public:
    RootBox(std::shared_ptr<Box> inner) : innerBox(inner) { UpdateMetrics(36.0f); }
    void UpdateMetrics(float fontSize) {
        padLeft = fontSize * (14.0f / 36.0f); padTop = fontSize * (3.0f / 36.0f); padRight = fontSize * (2.0f / 36.0f);
        float minInternalHeight = fontSize * (26.0f / 36.0f); float minInternalDepth = fontSize * (8.0f / 36.0f);
        actualOpeningHeight = std::max(innerBox->height, minInternalHeight); actualOpeningDepth = std::max(innerBox->depth, minInternalDepth);
        width = padLeft + innerBox->width + padRight; height = actualOpeningHeight + padTop; depth = actualOpeningDepth + 2.0f;
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift; innerBox->Draw(context, x + padLeft, currentY);
        float p1x = x; float p1y = currentY - actualOpeningHeight * 0.3f;
        float p2x = x + padLeft * 0.45f; float p2y = currentY + actualOpeningDepth + 2.0f;
        float p3x = x + padLeft - 2.0f; float p3y = currentY - actualOpeningHeight - padTop + 1.0f;
        float p4x = x + width; float p4y = p3y;
        context->DrawLine(p1x, p1y, p2x, p2y, 1.5f); context->DrawLine(p2x, p2y, p3x, p3y, 1.5f); context->DrawLine(p3x, p3y, p4x, p4y, 1.5f);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        innerBox->SetFontSize(newSize, ctx); UpdateMetrics(newSize);
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        float currentY = y + shift;
        innerBox->MapCaret(ctx, x + padLeft, currentY, scale, true, metrics);
    }
};
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
        m_pDWriteFactory->CreateTextFormat(L"Cambria Math", nullptr, DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 32.0f, L"en-us", &m_pFormatNormal);
        m_pDWriteFactory->CreateTextFormat(L"Cambria Math", nullptr, DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_ITALIC, DWRITE_FONT_STRETCH_NORMAL, 32.0f, L"en-us", &m_pFormatItalic);
    }
    void MeasureGlyph(const std::wstring& text, float fontSize, bool isItalic, float& outWidth, float& outHeight, float& outDepth) override {
        ComPtr<IDWriteTextLayout> pLayout;
        IDWriteTextFormat* format = isItalic ? m_pFormatItalic.Get() : m_pFormatNormal.Get();
        m_pDWriteFactory->CreateTextLayout(text.c_str(), static_cast<UINT32>(text.length()), format, 10000.0f, 10000.0f, &pLayout);
        DWRITE_TEXT_RANGE range = { 0, static_cast<UINT32>(text.length()) };
        pLayout->SetFontSize(fontSize, range);
        DWRITE_TEXT_METRICS metrics;
        pLayout->GetMetrics(&metrics);
        outWidth = metrics.widthIncludingTrailingWhitespace;
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
        DWRITE_LINE_METRICS lineMetrics;
        UINT32 actualLineCount;
        pLayout->GetLineMetrics(&lineMetrics, 1, &actualLineCount);
        m_pRT->DrawTextLayout(D2D1::Point2F(x, y - lineMetrics.baseline), pLayout.Get(), m_pBrush.Get());
    }
    void DrawLine(float x1, float y1, float x2, float y2, float thickness) override {
        m_pRT->DrawLine(D2D1::Point2F(x1, y1), D2D1::Point2F(x2, y2), m_pBrush.Get(), thickness);
    }
    void FillCircle(float x, float y, float r) override {
        D2D1_ELLIPSE ellipse = D2D1::Ellipse(D2D1::Point2F(x, y), r, r);
        m_pRT->FillEllipse(ellipse, m_pBrush.Get());
    }
    void SetTextColor(float r, float g, float b, float a = 1.0f) override {
        m_pBrush->SetColor(D2D1::ColorF(r, g, b, a));
    }
};
ComPtr<ID2D1Factory7> g_pD2DFactory;
ComPtr<IDWriteFactory3> g_pDWriteFactory;
ComPtr<ID2D1HwndRenderTarget> g_pRenderTarget;
std::shared_ptr<Box> g_pLayoutTree;
std::wstring g_inputText = L"";
int g_cursorPos = 0;
int g_selectionStart = 0;
bool g_caretVisible = true;
const UINT CARET_TIMER_ID = 1;
void DeleteSelection() {
    if (g_cursorPos == g_selectionStart) return;
    int startIdx = std::min(g_cursorPos, g_selectionStart);
    int endIdx = std::max(g_cursorPos, g_selectionStart);
    g_inputText.erase(startIdx, endIdx - startIdx);
    g_cursorPos = startIdx;
    g_selectionStart = startIdx;
}
void CopyToClipboard(HWND hWnd, const std::wstring& text) {
    if (text.empty()) return;
    if (!OpenClipboard(hWnd)) return;
    EmptyClipboard();
    size_t size = (text.length() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
    if (hMem) {
        void* ptr = GlobalLock(hMem);
        if (ptr) {
            memcpy(ptr, text.c_str(), size);
            GlobalUnlock(hMem);
            SetClipboardData(CF_UNICODETEXT, hMem);
        }
        else {
            GlobalFree(hMem);
        }
    }
    CloseClipboard();
}
std::wstring PasteFromClipboard(HWND hWnd) {
    std::wstring result = L"";
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) return result;
    if (!OpenClipboard(hWnd)) return result;
    HGLOBAL hMem = GetClipboardData(CF_UNICODETEXT);
    if (hMem) {
        const wchar_t* ptr = static_cast<const wchar_t*>(GlobalLock(hMem));
        if (ptr) {
            result = ptr;
            GlobalUnlock(hMem);
        }
    }
    CloseClipboard();
    return result;
}
float GetIcp(wchar_t op) {
    if (op == L'(') return 5.0f;
    if (op == L'!') return 4.8f;
    if (op == L'√') return 4.5f;
    if (op == L'^') return 4.0f;
    if (op == L'_' || op == L'#') return 3.5f;
    if (op == L'@') return 3.0f;
    if (op == L'/') return 2.0f;
    if (op == L'+' || op == L'-') return 1.0f;
    return 0.0f;
}
float GetIsp(wchar_t op) {
    if (op == L'(') return 0.0f;
    if (op == L'!') return 4.8f;
    if (op == L'√') return 2.5f;
    if (op == L'^') return 2.5f;
    if (op == L'_' || op == L'#') return 3.5f;
    if (op == L'@') return 3.0f;
    if (op == L'/') return 2.0f;
    if (op == L'+' || op == L'-') return 1.0f;
    return 0.0f;
}
struct OpToken {
    wchar_t ch;
    int idx;
};
bool ApplyOperator(std::stack<OpToken>& opStack, std::stack<std::shared_ptr<Box>>& boxStack, IMathRendererContext* ctx) {
    if (opStack.empty()) return false;
    OpToken opInfo = opStack.top();
    wchar_t op = opInfo.ch;
    if (op == L'√') {
        if (boxStack.empty()) return false;
        opStack.pop();
        std::shared_ptr<Box> inner = boxStack.top(); boxStack.pop();
        boxStack.push(std::make_shared<RootBox>(inner));
        return true;
    }
    if (op == L'!') {
        if (boxStack.empty()) return false;
        opStack.pop();
        std::shared_ptr<Box> inner = boxStack.top(); boxStack.pop();
        auto horiz = std::make_shared<HorizontalBox>();
        horiz->Add(inner);
        auto exclBox = std::make_shared<CharBox>(L"!", 36.0f, false, ctx);
        exclBox->srcStart = opInfo.idx;
        exclBox->srcEnd = opInfo.idx + 1;
        horiz->Add(exclBox);
        boxStack.push(horiz);
        return true;
    }
    if (op == L'_' || op == L'#') {
        if (boxStack.empty()) return false;
        opStack.pop();
        std::shared_ptr<Box> inner = boxStack.top(); boxStack.pop();
        auto horiz = std::make_shared<HorizontalBox>();
        wchar_t opChar = (op == L'_') ? L'\x2212' : L'+';
        auto signBox = std::make_shared<CharBox>(std::wstring(1, opChar), 36.0f, false, ctx);
        signBox->srcStart = opInfo.idx;
        signBox->srcEnd = opInfo.idx + 1;
        horiz->Add(signBox);
        horiz->Add(inner);
        boxStack.push(horiz);
        return true;
    }
    if (boxStack.size() < 2) return false;
    opStack.pop();
    std::shared_ptr<Box> right = boxStack.top(); boxStack.pop();
    std::shared_ptr<Box> left = boxStack.top(); boxStack.pop();
    std::shared_ptr<Box> result;
    float normalSize = 36.0f;
    float scriptSize = 22.0f;
    switch (op) {
    case L'@':
    {
        auto horiz = std::make_shared<HorizontalBox>();
        horiz->Add(left);
        horiz->Add(right);
        result = horiz;
        break;
    }
    case L'^':
        right->SetFontSize(scriptSize, ctx);
        result = std::make_shared<ScriptBox>(left, right, nullptr);
        break;
    case L'/':
        result = std::make_shared<FractionBox>(left, right);
        break;
    case L'+':
    case L'-':
    {
        auto horiz = std::make_shared<HorizontalBox>();
        horiz->Add(left);
        horiz->Add(std::make_shared<SpaceBox>(8.0f));
        wchar_t opChar = (op == L'-') ? L'\x2212' : L'+';
        auto opBox = std::make_shared<CharBox>(std::wstring(1, opChar), normalSize, false, ctx);
        opBox->srcStart = opInfo.idx;
        opBox->srcEnd = opInfo.idx + 1;
        horiz->Add(opBox);
        horiz->Add(std::make_shared<SpaceBox>(8.0f));
        horiz->Add(right);
        result = horiz;
    }
    break;
    }
    if (result) {
        boxStack.push(result);
        return true;
    }
    return false;
}
bool IsFunctionName(const std::wstring& text, size_t pos, std::wstring& outName) {
    static const std::vector<std::wstring> funcNames = {
        L"arcsin", L"arccos", L"arctan",
        L"sinh", L"cosh", L"tanh", L"asin", L"acos", L"atan",
        L"sin", L"cos", L"tan", L"log", L"lim", L"exp", L"max", L"min", L"det",
        L"ln"
    };
    for (const auto& name : funcNames) {
        if (pos + name.length() <= text.length() &&
            text.substr(pos, name.length()) == name) {
            outName = name;
            return true;
        }
    }
    return false;
}
BigFloat EvalExpr(const wchar_t*& p);
BigFloat EvalFactor(const wchar_t*& p) {
    while (*p == L' ' || *p == L'\x200B') p++;
    if (*p == L'-') { p++; return -EvalFactor(p); }
    if (*p == L'+') { p++; return EvalFactor(p); }
    BigFloat val = 0.0;
    if (*p == L'(') {
        p++; val = EvalExpr(p);
        if (*p == L')') p++;
    }
    else if (wcsncmp(p, L"sin", 3) == 0) { p += 3; val = sin(EvalFactor(p)); }
    else if (wcsncmp(p, L"cos", 3) == 0) { p += 3; val = cos(EvalFactor(p)); }
    else if (wcsncmp(p, L"log", 3) == 0) { p += 3; val = log10(EvalFactor(p)); }
    else if (*p == L'√') { p++; val = sqrt(EvalFactor(p)); }
    else if (iswdigit(*p) || *p == L'.') {
        std::string numStr;
        while (iswdigit(*p) || *p == L'.') {
            numStr += static_cast<char>(*p);
            p++;
        }
        val = BigFloat(numStr);
    }
    else if (iswalpha(*p)) {
        p++;
        val = std::numeric_limits<BigFloat>::quiet_NaN();
    }
    while (*p == L' ' || *p == L'\x200B') p++;
    while (*p == L'!') {
        p++;
        if (val > 1000.0 || val < 0.0) {
            return std::numeric_limits<BigFloat>::quiet_NaN();
        }
        val = boost::math::tgamma(val + 1);
        while (*p == L' ' || *p == L'\x200B') p++;
    }
    if (*p == L'^') {
        p++;
        BigFloat exponent = EvalFactor(p);
        if (abs(exponent) > 10000.0) {
            val = std::numeric_limits<BigFloat>::quiet_NaN();
        }
        else {
            val = pow(val, exponent);
        }
    }
    return val;
}
BigFloat EvalTerm(const wchar_t*& p) {
    BigFloat val = EvalFactor(p);
    while (true) {
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'*') { p++; val *= EvalFactor(p); }
        else if (*p == L'/') { p++; val /= EvalFactor(p); }
        else if (iswdigit(*p) || iswalpha(*p) || *p == L'(' || *p == L'√' || wcsncmp(p, L"sin", 3) == 0 || wcsncmp(p, L"cos", 3) == 0) {
            val *= EvalFactor(p);
        }
        else break;
    }
    return val;
}
BigFloat EvalExpr(const wchar_t*& p) {
    while (*p == L' ' || *p == L'\x200B') p++;
    BigFloat val = EvalTerm(p);
    while (true) {
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'+') { p++; val += EvalTerm(p); }
        else if (*p == L'-') { p++; val -= EvalTerm(p); }
        else break;
    }
    return val;
}
std::wstring FormatRepeatingDecimal(const std::wstring& numStr) {
    if (numStr.find(L'e') != std::wstring::npos || numStr.find(L'E') != std::wstring::npos) return numStr;
    size_t dotPos = numStr.find(L'.');
    if (dotPos == std::wstring::npos) return numStr;
    std::wstring intPart = numStr.substr(0, dotPos + 1);
    std::wstring frac = numStr.substr(dotPos + 1);
    if (frac.length() > 20) {
        std::wstring safeFrac = frac.substr(0, frac.length() - 5);
        int maxLen = static_cast<int>(safeFrac.length());
        for (int start = 0; start < 15; ++start) {
            if (start >= maxLen) break;
            for (int p = 1; p <= (maxLen - start) / 2; ++p) {
                std::wstring pattern = safeFrac.substr(start, p);
                if (pattern.find_first_not_of(L'9') == std::wstring::npos) continue;
                bool match = true;
                int repeats = 0;
                for (int i = start + p; i + p <= maxLen; i += p) {
                    if (safeFrac.substr(i, p) != pattern) {
                        match = false; break;
                    }
                    repeats++;
                }
                if (match && repeats >= 3) {
                    std::wstring result = intPart + frac.substr(0, start);
                    if (p == 1) {
                        result += L"\x02" + pattern;
                    }
                    else {
                        result += L"\x02" + pattern.substr(0, 1) +
                            pattern.substr(1, p - 2) +
                            L"\x02" + pattern.substr(p - 1, 1);
                    }
                    return result;
                }
            }
        }
    }
    return numStr;
}
std::wstring CalculateResult(std::wstring text) {
    text.erase(std::remove(text.begin(), text.end(), L'\x200B'), text.end());
    if (text.empty()) return L"";
    if (text.find(L'=') != std::wstring::npos) return L"";
    bool hasOperator = false;
    for (wchar_t ch : text) {
        if (ch == L'+' || ch == L'-' || ch == L'*' || ch == L'/' ||
            ch == L'^' || ch == L'√' || ch == L'(' || ch == L')' ||
            ch == L'!' || iswalpha(ch)) {
            hasOperator = true;
            break;
        }
    }
    if (!hasOperator) return L"";
    wchar_t last = text.back();
    if (last == L'+' || last == L'-' || last == L'*' || last == L'/' || last == L'^' || last == L'√' || last == L'(') {
        return L"";
    }
    const wchar_t* p = text.c_str();
    BigFloat result;
    try {
        result = EvalExpr(p);
    }
    catch (...) {
        return L"";
    }
    while (*p == L' ' || *p == L'\x200B') p++;
    if (*p != L'\0') return L"";
    if (boost::math::isnan(result) || boost::math::isinf(result)) return L"";
    std::string resStr;
    bool formatted = false;
    for (int prec = 150; prec >= 0; prec -= 10) {
        try {
            std::ostringstream oss;
            oss << std::fixed << std::setprecision(prec) << result;
            resStr = oss.str();
            formatted = true;
            break;
        }
        catch (...) {
        }
    }

    if (!formatted) {
        std::ostringstream oss;
        oss << result;
        resStr = oss.str();
    }
    if (resStr.find('e') == std::string::npos && resStr.find('E') == std::string::npos) {
        if (resStr.find('.') != std::string::npos) {
            resStr.erase(resStr.find_last_not_of('0') + 1, std::string::npos);
            if (!resStr.empty() && resStr.back() == '.') {
                resStr.pop_back();
            }
        }
    }
    std::wstring wresStr(resStr.begin(), resStr.end());
    wresStr = FormatRepeatingDecimal(wresStr);
    return wresStr;
}
std::shared_ptr<Box> ParseMathText(const std::wstring& text, IMathRendererContext* ctx) {
    if (text.empty()) return std::make_shared<HorizontalBox>();
    std::stack<std::shared_ptr<Box>> boxStack;
    std::stack<OpToken> opStack;
    std::stack<size_t> parenBoxStackSizes;
    float normalSize = 36.0f;
    bool expectOperand = true;
    auto EvaluateTop = [&](std::stack<OpToken>& ops, std::stack<std::shared_ptr<Box>>& boxes, int currentIndex) {
        bool pushedDummy = false;
        if (expectOperand) {
            auto dummy = std::make_shared<HorizontalBox>();
            dummy->srcStart = currentIndex;
            dummy->srcEnd = currentIndex;
            boxes.push(dummy);
            expectOperand = false;
            pushedDummy = true;
        }
        if (!ApplyOperator(ops, boxes, ctx)) {
            if (pushedDummy) boxes.pop();
            OpToken failedOp = ops.top();
            ops.pop();
            if (failedOp.ch != L'@') {
                wchar_t fallbackCh = failedOp.ch;
                if (fallbackCh == L'_') fallbackCh = L'-';
                if (fallbackCh == L'#') fallbackCh = L'+';
                auto charBox = std::make_shared<CharBox>(std::wstring(1, fallbackCh), normalSize, false, ctx);
                charBox->srcStart = failedOp.idx;
                charBox->srcEnd = failedOp.idx + 1;
                boxes.push(charBox);
            }
        }
        };
    for (size_t i = 0; i < text.length(); ) {
        wchar_t ch = text[i];
        if (ch == L' ') {
            if (!expectOperand) {
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= 2.0f) EvaluateTop(opStack, boxStack, static_cast<int>(i));
            }
            i++; continue;
        }
        else if (ch == L'\x200B') {
            if (expectOperand) {
                auto dummy = std::make_shared<HorizontalBox>();
                dummy->srcStart = (int)i; dummy->srcEnd = (int)i;
                boxStack.push(dummy);
                expectOperand = false;
            }
            while (!opStack.empty() && GetIsp(opStack.top().ch) >= 2.0f) EvaluateTop(opStack, boxStack, static_cast<int>(i));
            OpToken op = { L'@', -1 };
            while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
            opStack.push(op);
            auto zwspBox = std::make_shared<SpaceBox>(0.0f);
            zwspBox->srcStart = (int)i; zwspBox->srcEnd = (int)(i + 1);
            boxStack.push(zwspBox);
            expectOperand = false; i++;
            continue;
        }
        else if (ch == L'\x02') {
            if (!expectOperand) {
                OpToken op = { L'@', -1 };
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push(op);
            }
            i++;
            if (i < text.length()) {
                wchar_t targetCh = text[i];
                auto charBox = std::make_shared<CharBox>(std::wstring(1, targetCh), normalSize, false, ctx);
                charBox->srcStart = (int)i; charBox->srcEnd = (int)(i + 1);
                boxStack.push(std::make_shared<DotBox>(charBox));
                expectOperand = false;
                i++;
            }
            continue;
        }
        if (iswdigit(ch) || ch == L'.') {
            if (!expectOperand) {
                OpToken op = { L'@', -1 };
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push(op);
            }
            size_t startI = i;
            std::wstring numStr = L"";
            while (i < text.length() && (iswdigit(text[i]) || text[i] == L'.')) {
                numStr += text[i]; i++;
            }
            auto box = std::make_shared<CharBox>(numStr, normalSize, false, ctx);
            box->srcStart = (int)startI;
            box->srcEnd = (int)i;
            boxStack.push(box);
            expectOperand = false;
        }
        else if (iswalpha(ch)) {
            std::wstring funcName;
            if (IsFunctionName(text, i, funcName)) {
                if (!expectOperand) {
                    OpToken op = { L'@', -1 };
                    while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                    opStack.push(op);
                }
                auto box = std::make_shared<CharBox>(funcName, normalSize, false, ctx);
                box->srcStart = (int)i;
                box->srcEnd = (int)(i + funcName.length());
                boxStack.push(box);
                expectOperand = false;
                i += funcName.length();
            }
            else {
                if (!expectOperand) {
                    OpToken op = { L'@', -1 };
                    while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                    opStack.push(op);
                }
                auto box = std::make_shared<CharBox>(std::wstring(1, ch), normalSize, true, ctx);
                box->srcStart = (int)i;
                box->srcEnd = (int)(i + 1);
                boxStack.push(box);
                expectOperand = false;
                i++;
            }
        }
        else if (ch == L'√') {
            if (!expectOperand) {
                OpToken op = { L'@', -1 };
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push(op);
            }
            opStack.push({ ch, (int)i });
            expectOperand = true; i++;
        }
        else if (ch == L'!') {
            if (expectOperand) {
                auto dummy = std::make_shared<HorizontalBox>();
                dummy->srcStart = (int)i;
                dummy->srcEnd = (int)i;
                boxStack.push(dummy);
            }
            else {
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
            }
            opStack.push({ ch, (int)i });
            expectOperand = false;
            i++;
        }
        else if (ch == L'(') {
            if (!expectOperand) {
                OpToken op = { L'@', -1 };
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push(op);
            }
            opStack.push({ ch, (int)i });
            parenBoxStackSizes.push(boxStack.size());
            expectOperand = true; i++;
        }
        else if (ch == L')') {
            while (!opStack.empty() && opStack.top().ch != L'(') {
                EvaluateTop(opStack, boxStack, static_cast<int>(i));
            }
            if (!opStack.empty() && opStack.top().ch == L'(') {
                int leftParenIdx = opStack.top().idx;
                opStack.pop();
                size_t baseSize = parenBoxStackSizes.top();
                parenBoxStackSizes.pop();
                auto wrap = std::make_shared<HorizontalBox>();
                auto lBox = std::make_shared<CharBox>(L"(", normalSize, false, ctx);
                lBox->srcStart = leftParenIdx; lBox->srcEnd = leftParenIdx + 1;
                wrap->Add(lBox);
                if (boxStack.size() > baseSize) {
                    std::vector<std::shared_ptr<Box>> inners;
                    while (boxStack.size() > baseSize) {
                        inners.push_back(boxStack.top());
                        boxStack.pop();
                    }
                    for (auto it = inners.rbegin(); it != inners.rend(); ++it) {
                        wrap->Add(*it);
                    }
                }
                auto rBox = std::make_shared<CharBox>(L")", normalSize, false, ctx);
                rBox->srcStart = (int)i; rBox->srcEnd = (int)(i + 1);
                wrap->Add(rBox);
                boxStack.push(wrap);
            }
            else {
                if (!expectOperand) {
                    OpToken op = { L'@', -1 };
                    while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                    opStack.push(op);
                }
                auto rBox = std::make_shared<CharBox>(L")", normalSize, false, ctx);
                rBox->srcStart = (int)i; rBox->srcEnd = (int)(i + 1);
                boxStack.push(rBox);
            }
            expectOperand = false; i++;
        }
        else if (ch == L'+' || ch == L'-') {
            if (expectOperand) {
                wchar_t unaryOp = (ch == L'-') ? L'_' : L'#';
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(unaryOp)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push({ unaryOp, (int)i });
                expectOperand = true;
            }
            else {
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push({ ch, (int)i });
                expectOperand = true;
            }
            i++;
        }
        else if (ch == L'^' || ch == L'/') {
            while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
            opStack.push({ ch, (int)i });
            expectOperand = true; i++;
        }
        else {
            if (!expectOperand) {
                OpToken op = { L'@', -1 };
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(op.ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
                opStack.push(op);
            }
            auto box = std::make_shared<CharBox>(std::wstring(1, ch), normalSize, false, ctx);
            box->srcStart = (int)i; box->srcEnd = (int)(i + 1);
            boxStack.push(box);
            expectOperand = false; i++;
        }
    }
    while (!opStack.empty()) {
        OpToken opInfo = opStack.top();
        if (opInfo.ch == L'(') {
            opStack.pop();
            size_t baseSize = parenBoxStackSizes.top();
            parenBoxStackSizes.pop();
            auto wrap = std::make_shared<HorizontalBox>();
            auto lBox = std::make_shared<CharBox>(L"(", normalSize, false, ctx);
            lBox->srcStart = opInfo.idx; lBox->srcEnd = opInfo.idx + 1;
            wrap->Add(lBox);
            if (boxStack.size() > baseSize) {
                std::vector<std::shared_ptr<Box>> inners;
                while (boxStack.size() > baseSize) {
                    inners.push_back(boxStack.top());
                    boxStack.pop();
                }
                for (auto it = inners.rbegin(); it != inners.rend(); ++it) {
                    wrap->Add(*it);
                }
            }
            boxStack.push(wrap);
            expectOperand = false;
            continue;
        }
        EvaluateTop(opStack, boxStack, static_cast<int>(text.length()));
    }
    if (boxStack.empty()) return std::make_shared<HorizontalBox>();
    if (boxStack.size() == 1) return boxStack.top();
    auto root = std::make_shared<HorizontalBox>();
    std::vector<std::shared_ptr<Box>> leftover;
    while (!boxStack.empty()) { leftover.push_back(boxStack.top()); boxStack.pop(); }
    for (auto it = leftover.rbegin(); it != leftover.rend(); ++it) root->Add(*it);
    return root;
}
int GetCursorPosFromMouse(HWND hWnd, int mouseX, int mouseY) {
    if (g_caretMetrics.empty()) return 0;
    RECT clientRect;
    GetClientRect(hWnd, &clientRect);
    float startX = 50.0f;
    float startY = (clientRect.bottom - clientRect.top) / 2.0f;
    int bestIndex = 0;
    float minDistanceSq = -1.0f;
    for (size_t i = 0; i < g_caretMetrics.size(); ++i) {
        if (g_caretMetrics[i].x < 0.0f) continue;
        float cx = startX + g_caretMetrics[i].x;
        float cy = startY + g_caretMetrics[i].y - 12.0f;
        float dx = mouseX - cx;
        float dy = mouseY - cy;
        float distSq = dx * dx + dy * dy * 5.0f;
        if (minDistanceSq < 0.0f || distSq < minDistanceSq) {
            minDistanceSq = distSq;
            bestIndex = static_cast<int>(i);
        }
    }
    return bestIndex;
}
bool IsSystemDarkMode() {
    DWORD value = 1;
    DWORD size = sizeof(value);
    LSTATUS status = RegGetValueW(
        HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
        L"AppsUseLightTheme",
        RRF_RT_REG_DWORD,
        nullptr,
        &value,
        &size
    );
    if (status == ERROR_SUCCESS) {
        return value == 0;
    }
    return false;
}
void ApplyTitleBarTheme(HWND hWnd, bool isDark) {
    BOOL dark = isDark ? TRUE : FALSE;
    DwmSetWindowAttribute(hWnd, 20, &dark, sizeof(dark));
}
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        g_isDarkMode = IsSystemDarkMode();
        ApplyTitleBarTheme(hWnd, g_isDarkMode);
        SetTimer(hWnd, CARET_TIMER_ID, 500, nullptr);
        break;
    case WM_SETTINGCHANGE:
        if (lParam != 0) {
            LPCWSTR str = reinterpret_cast<LPCWSTR>(lParam);
            if (wcscmp(str, L"ImmersiveColorSet") == 0) {
                g_isDarkMode = IsSystemDarkMode();
                ApplyTitleBarTheme(hWnd, g_isDarkMode);
                InvalidateRect(hWnd, nullptr, FALSE);
            }
        }
        break;
    case WM_TIMER:
        if (wParam == CARET_TIMER_ID) {
            g_caretVisible = !g_caretVisible;
            InvalidateRect(hWnd, nullptr, FALSE);
        }
        break;
    case WM_LBUTTONDOWN:
    {
        int xPos = GET_X_LPARAM(lParam);
        int yPos = GET_Y_LPARAM(lParam);
        int newPos = GetCursorPosFromMouse(hWnd, xPos, yPos);
        g_cursorPos = newPos;
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0) {
            g_selectionStart = g_cursorPos;
        }
        SetCapture(hWnd);
        g_isDragging = true;
        g_caretVisible = true;
        InvalidateRect(hWnd, nullptr, FALSE);
    }
    break;
    case WM_LBUTTONDBLCLK:
    {
        if (g_inputText.empty()) break;
        int xPos = GET_X_LPARAM(lParam);
        int yPos = GET_Y_LPARAM(lParam);
        int clickPos = GetCursorPosFromMouse(hWnd, xPos, yPos);
        auto isWordChar = [](wchar_t c) {
            return iswdigit(c) || iswalpha(c) || c == L'.';
            };
        int leftIdx = clickPos;
        int rightIdx = clickPos;
        bool isInsideWord = false;
        if (clickPos < g_inputText.length() && isWordChar(g_inputText[clickPos])) {
            isInsideWord = true;
        }
        else if (clickPos > 0 && isWordChar(g_inputText[clickPos - 1])) {
            isInsideWord = true;
        }
        if (isInsideWord) {
            while (leftIdx > 0 && isWordChar(g_inputText[leftIdx - 1])) {
                leftIdx--;
            }
            while (rightIdx < g_inputText.length() && isWordChar(g_inputText[rightIdx])) {
                rightIdx++;
            }
        }
        else {
            if (clickPos < g_inputText.length()) {
                rightIdx = clickPos + 1;
            }
            else if (clickPos > 0) {
                leftIdx = clickPos - 1;
            }
        }
        g_selectionStart = leftIdx;
        g_cursorPos = rightIdx;
        g_caretVisible = true;
        InvalidateRect(hWnd, nullptr, FALSE);
    }
    break;
    case WM_MOUSEMOVE:
    {
        if (g_isDragging) {
            int xPos = GET_X_LPARAM(lParam);
            int yPos = GET_Y_LPARAM(lParam);
            int newPos = GetCursorPosFromMouse(hWnd, xPos, yPos);
            if (g_cursorPos != newPos) {
                g_cursorPos = newPos;
                g_caretVisible = true;
                InvalidateRect(hWnd, nullptr, FALSE);
            }
        }
    }
    break;
    case WM_LBUTTONUP:
    {
        if (g_isDragging) {
            ReleaseCapture();
            g_isDragging = false;
        }
    }
    break;
    case WM_KEYDOWN:
    {
        bool shiftDown = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        auto executePaste = [&]() {
            std::wstring pastedText = PasteFromClipboard(hWnd);
            if (!pastedText.empty()) {
                if (g_cursorPos != g_selectionStart) DeleteSelection();
                g_inputText.insert(g_cursorPos, pastedText);
                g_cursorPos += static_cast<int>(pastedText.length());
                g_selectionStart = g_cursorPos;
            }
            };
        switch (wParam) {
        case VK_RETURN:
            g_resultText = CalculateResult(g_inputText);
            g_pResultTree.reset();
            InvalidateRect(hWnd, nullptr, FALSE);
            break;
        case VK_LEFT:
            if (g_cursorPos > 0) g_cursorPos--;
            if (!shiftDown) g_selectionStart = g_cursorPos;
            break;
        case VK_RIGHT:
            if (g_cursorPos < (int)g_inputText.length()) {
                g_cursorPos++;
            }
            else {
                if (!g_caretMetrics.empty() && g_caretMetrics[g_cursorPos].isActive) {
                    g_inputText.insert(g_cursorPos, 1, L'\x200B');
                    g_cursorPos++;
                }
            }
            if (!shiftDown) g_selectionStart = g_cursorPos;
            break;
        case VK_HOME:
            g_cursorPos = 0;
            if (!shiftDown) g_selectionStart = g_cursorPos;
            break;
        case VK_END:
            g_cursorPos = (int)g_inputText.length();
            if (!shiftDown) g_selectionStart = g_cursorPos;
            break;
        case VK_BACK:
            if (g_cursorPos != g_selectionStart) {
                DeleteSelection();
            }
            else if (g_cursorPos > 0) {
                g_inputText.erase(g_cursorPos - 1, 1);
                g_cursorPos--;
                g_selectionStart = g_cursorPos;
            }
            break;
        case VK_DELETE:
            if (g_cursorPos != g_selectionStart) {
                DeleteSelection();
            }
            else if (g_cursorPos < (int)g_inputText.length()) {
                g_inputText.erase(g_cursorPos, 1);
            }
            break;
        case 'R':
            if (ctrlDown) {
                if (g_cursorPos != g_selectionStart) DeleteSelection();
                g_inputText.insert(g_cursorPos, L"√");
                g_cursorPos++;
                g_selectionStart = g_cursorPos;
                g_pLayoutTree.reset();
                g_caretVisible = true;
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            break;
        case 'A':
            if (ctrlDown) {
                g_selectionStart = 0;
                g_cursorPos = static_cast<int>(g_inputText.length());
                g_caretVisible = true;
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            break;
        case 'C':
            if (ctrlDown && g_cursorPos != g_selectionStart) {
                int startIdx = std::min(g_cursorPos, g_selectionStart);
                int endIdx = std::max(g_cursorPos, g_selectionStart);
                CopyToClipboard(hWnd, g_inputText.substr(startIdx, endIdx - startIdx));
            }
            break;
        case 'X':
            if (ctrlDown && g_cursorPos != g_selectionStart) {
                int startIdx = std::min(g_cursorPos, g_selectionStart);
                int endIdx = std::max(g_cursorPos, g_selectionStart);
                CopyToClipboard(hWnd, g_inputText.substr(startIdx, endIdx - startIdx));
                DeleteSelection();
            }
            break;
        case 'V':
            if (ctrlDown) {
                executePaste();
            }
            break;
        case VK_INSERT:
            if (shiftDown) {
                executePaste();
            }
            break;
        }
        g_resultText = CalculateResult(g_inputText);
        g_pResultTree.reset();
        g_pLayoutTree.reset();
        g_caretVisible = true;
        InvalidateRect(hWnd, nullptr, FALSE);
    }
    break;
    case WM_CHAR:
    {
        wchar_t ch = static_cast<wchar_t>(wParam);
        if (ch >= 32) {
            if (g_cursorPos != g_selectionStart) {
                DeleteSelection();
            }
            bool shouldAutoCompleteFraction = false;
            if (ch == L'/') {
                if (g_cursorPos == 0) {
                    shouldAutoCompleteFraction = true;
                }
                else {
                    wchar_t prevChar = g_inputText[g_cursorPos - 1];
                    if (prevChar == L'(' || prevChar == L'^' || prevChar == L'/' ||
                        prevChar == L'+' || prevChar == L'-' || prevChar == L'=' ||
                        prevChar == L' ' || prevChar == L'\x200B') {
                        shouldAutoCompleteFraction = true;
                    }
                }
            }
            if (shouldAutoCompleteFraction) {
                g_inputText.insert(g_cursorPos, L"1/");
                g_cursorPos += 2;
            }
            else {
                g_inputText.insert(g_cursorPos, 1, ch);
                g_cursorPos++;
            }
            g_selectionStart = g_cursorPos;
            g_resultText = CalculateResult(g_inputText);
            g_pResultTree.reset();
            g_pLayoutTree.reset();
            g_caretVisible = true;
            InvalidateRect(hWnd, nullptr, FALSE);
        }
    }
    break;
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        BeginPaint(hWnd, &ps);
        if (!g_pRenderTarget) {
            RECT rc;
            GetClientRect(hWnd, &rc);
            D2D1_SIZE_U size = D2D1::SizeU(rc.right - rc.left, rc.bottom - rc.top);
            g_pD2DFactory->CreateHwndRenderTarget(D2D1::RenderTargetProperties(), D2D1::HwndRenderTargetProperties(hWnd, size), &g_pRenderTarget);
        }
        g_pRenderTarget->BeginDraw();

        // ★修正: 背景色をモードによって切り替え (ダークモード時は VSCode のような深いグレーに)
        D2D1_COLOR_F bgColor = g_isDarkMode ? D2D1::ColorF(0x1E1E1E) : D2D1::ColorF(D2D1::ColorF::White);
        g_pRenderTarget->Clear(bgColor);

        Direct2DContext ctx(g_pRenderTarget.Get(), g_pDWriteFactory.Get());

        // ★修正: 基本の文字色を切り替え
        D2D1_COLOR_F textColor = g_isDarkMode ? D2D1::ColorF(0xD4D4D4) : D2D1::ColorF(D2D1::ColorF::Black);
        ctx.SetTextColor(textColor.r, textColor.g, textColor.b, textColor.a);

        if (!g_pLayoutTree) g_pLayoutTree = ParseMathText(g_inputText, &ctx);
        int len = static_cast<int>(g_inputText.length());
        std::vector<CaretMetrics> metrics(len + 1);
        g_pLayoutTree->MapCaret(&ctx, 0.0f, 0.0f, 1.0f, false, metrics);
        if (metrics[0].x < 0.0f) {
            metrics[0].x = 0.0f; metrics[0].y = 0.0f;
            metrics[0].scale = 1.0f; metrics[0].isActive = false;
        }
        for (int i = 1; i <= len; i++) {
            if (metrics[i].x < 0.0f) {
                metrics[i].x = metrics[i - 1].x; metrics[i].y = metrics[i - 1].y;
                metrics[i].scale = metrics[i - 1].scale; metrics[i].isActive = metrics[i - 1].isActive;
            }
        }
        for (int i = len - 1; i >= 0; i--) {
            if (metrics[i].x < 0.0f) {
                metrics[i].x = metrics[i + 1].x; metrics[i].y = metrics[i + 1].y;
                metrics[i].scale = metrics[i + 1].scale; metrics[i].isActive = metrics[i + 1].isActive;
            }
        }
        g_caretMetrics = metrics;
        RECT clientRect;
        GetClientRect(hWnd, &clientRect);
        float startX = 50.0f;
        float startY = (clientRect.bottom - clientRect.top) / 2.0f;
        if (g_cursorPos != g_selectionStart) {
            int startIdx = std::min(g_cursorPos, g_selectionStart);
            int endIdx = std::max(g_cursorPos, g_selectionStart);
            ComPtr<ID2D1SolidColorBrush> highlightBrush;
            D2D1_COLOR_F selColor = g_isDarkMode ? D2D1::ColorF(0.15f, 0.35f, 0.6f, 0.6f) : D2D1::ColorF(0.0f, 0.4f, 1.0f, 0.2f);
            g_pRenderTarget->CreateSolidColorBrush(selColor, &highlightBrush);
            float currentLeftX = -10000.0f;
            float currentRightX = -10000.0f;
            float currentY = -10000.0f;
            float currentScale = -1.0f;
            auto drawCurrentRect = [&]() {
                if (currentLeftX != -10000.0f && currentRightX > currentLeftX) {
                    float baseTop = -30.0f;
                    float baseBottom = 12.0f;
                    float cTop = startY + currentY + (baseTop * currentScale);
                    float cBot = startY + currentY + (baseBottom * currentScale);
                    D2D1_RECT_F rect = D2D1::RectF(currentLeftX, cTop, currentRightX, cBot);
                    g_pRenderTarget->FillRectangle(&rect, highlightBrush.Get());
                }
                };
            for (int k = startIdx; k < endIdx; ++k) {
                float x1 = startX + g_caretMetrics[k].x;
                float y1 = g_caretMetrics[k].y;
                float scale1 = g_caretMetrics[k].scale;
                float x2 = startX + g_caretMetrics[k + 1].x;
                float y2 = g_caretMetrics[k + 1].y;
                float scale2 = g_caretMetrics[k + 1].scale;
                if (std::abs(y1 - y2) < 1.0f && std::abs(scale1 - scale2) < 0.01f && std::abs(x1 - x2) > 0.01f) {
                    float segLeft = std::min(x1, x2);
                    float segRight = std::max(x1, x2);
                    if (currentY != -10000.0f && std::abs(currentY - y1) < 1.0f && std::abs(currentScale - scale1) < 0.01f) {
                        currentLeftX = std::min(currentLeftX, segLeft);
                        currentRightX = std::max(currentRightX, segRight);
                    }
                    else {
                        drawCurrentRect();
                        currentLeftX = segLeft;
                        currentRightX = segRight;
                        currentY = y1;
                        currentScale = scale1;
                    }
                }
            }
            drawCurrentRect();
        }
        g_pLayoutTree->Draw(&ctx, startX, startY);
        if (!g_resultText.empty()) {
            if (g_isDarkMode) {
                ctx.SetTextColor(0.5f, 0.5f, 0.5f);
            }
            else {
                ctx.SetTextColor(0.6f, 0.6f, 0.6f);
            }            float gap = 20.0f;
            float eqW, eqH, eqD;
            ctx.MeasureGlyph(L"=", 36.0f, false, eqW, eqH, eqD);
            float equalX = startX + g_pLayoutTree->width + gap;
            ctx.DrawGlyph(L"=", equalX, startY, 36.0f, false);
            float resultX = equalX + eqW + gap;
            RECT rc;
            GetClientRect(hWnd, &rc);
            float maxResultWidth = rc.right - resultX - 20.0f;
            if (!g_pResultTree) {
                g_pResultTree = ParseMathText(g_resultText, &ctx);

                if (g_pResultTree->width > maxResultWidth) {
                    if (maxResultWidth <= 40.0f) {
                        g_pResultTree = ParseMathText(L"...", &ctx);
                    }
                    else {
                        std::wstring tempText = g_resultText;
                        while (tempText.length() > 0) {
                            tempText.pop_back();
                            if (!tempText.empty() && tempText.back() == L'\x02') {
                                tempText.pop_back();
                            }
                            auto tempTree = ParseMathText(tempText + L"...", &ctx);
                            if (tempTree->width <= maxResultWidth || tempText.empty()) {
                                g_pResultTree = tempTree;
                                break;
                            }
                        }
                    }
                }
            }
            g_pResultTree->Draw(&ctx, resultX, startY);
            ctx.SetTextColor(textColor.r, textColor.g, textColor.b, textColor.a);
        }
        if (g_caretVisible) {
            float caretX = startX + g_caretMetrics[g_cursorPos].x;
            float scale = g_caretMetrics[g_cursorPos].scale;
            float yOffset = g_caretMetrics[g_cursorPos].y;
            float baseTop = -26.0f;
            float baseBottom = 8.0f;
            float caretTop = startY + yOffset + (baseTop * scale);
            float caretBottom = startY + yOffset + (baseBottom * scale);
            ctx.DrawLine(caretX + 2.0f, caretTop, caretX + 2.0f, caretBottom, 2.0f);
        }
        g_pRenderTarget->EndDraw();
        EndPaint(hWnd, &ps);
    }
    break;
    case WM_SIZE:
    {
        if (g_pRenderTarget) {
            UINT width = LOWORD(lParam);
            UINT height = HIWORD(lParam);
            g_pRenderTarget->Resize(D2D1::SizeU(width, height));
        }
        g_pResultTree.reset();
        InvalidateRect(hWnd, nullptr, FALSE);
    }
    break;
    case WM_DESTROY:
        KillTimer(hWnd, CARET_TIMER_ID);
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
    wcex.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.lpszClassName = L"4mulaWindow";
    RegisterClassExW(&wcex);
    HWND hWnd = CreateWindowW(L"4mulaWindow", L"4mula", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, 800, 320, nullptr, nullptr, hInstance, nullptr);
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