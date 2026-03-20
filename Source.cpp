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
#include <cmath>
#include <algorithm>
#include <stack>

using namespace Microsoft::WRL;

class Box;
std::wstring g_resultText = L"";
std::shared_ptr<Box> g_pResultTree = nullptr;

class IMathRendererContext {
public:
    virtual ~IMathRendererContext() = default;
    virtual void DrawGlyph(const std::wstring& text, float x, float y, float fontSize, bool isItalic) = 0;
    virtual void DrawLine(float x1, float y1, float x2, float y2, float thickness) = 0;
    virtual void MeasureGlyph(const std::wstring& text, float fontSize, bool isItalic, float& outWidth, float& outHeight, float& outDepth) = 0;
};

class Box {
public:
    float width = 0.0f;
    float height = 0.0f;
    float depth = 0.0f;
    float shift = 0.0f;
    virtual ~Box() = default;
    virtual void Draw(IMathRendererContext* context, float x, float y) const = 0;
    virtual void SetFontSize(float newSize, IMathRendererContext* ctx) {}
    virtual float GetCaretX(bool isActive) const { return width; }
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
        context->DrawGlyph(character, x, y + shift, fontSize, isItalic); // 自身のshiftを適用
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        fontSize = newSize;
        ctx->MeasureGlyph(character, fontSize, isItalic, width, height, depth);
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
        float currentY = y + shift;
        for (const auto& child : children) {
            child->Draw(context, currentX, currentY);
            currentX += child->width;
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
    float GetCaretX(bool isActive) const override {
        if (children.empty()) return 0.0f;
        float cx = 0.0f;
        for (size_t i = 0; i < children.size() - 1; ++i) {
            cx += children[i]->width;
        }
        cx += children.back()->GetCaretX(isActive);
        return cx;
    }
};

class FractionBox : public Box {
    std::shared_ptr<Box> numerator;
    std::shared_ptr<Box> denominator;
    float paddingX = 4.0f;
    float mathAxisOffset = 12.0f;
    float gap = 5.0f;
public:
    FractionBox(std::shared_ptr<Box> num, std::shared_ptr<Box> den) : numerator(num), denominator(den) {
        width = std::max(num->width, den->width) + paddingX * 2.0f;
        height = mathAxisOffset + gap + num->depth + num->height;
        depth = (den->height + gap + den->depth) - mathAxisOffset;
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift;
        float mathAxis = currentY - mathAxisOffset;

        float numX = x + (width - numerator->width) / 2.0f;
        numerator->Draw(context, numX, mathAxis - gap - numerator->depth);

        float denX = x + (width - denominator->width) / 2.0f;
        denominator->Draw(context, denX, mathAxis + gap + denominator->height);

        context->DrawLine(x, mathAxis, x + width, mathAxis, 1.5f);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        numerator->SetFontSize(newSize, ctx);
        denominator->SetFontSize(newSize, ctx);
        paddingX = newSize * (4.0f / 36.0f);
        mathAxisOffset = newSize * (12.0f / 36.0f);
        gap = newSize * (5.0f / 36.0f);
        width = std::max(numerator->width, denominator->width) + paddingX * 2.0f;
        height = mathAxisOffset + gap + numerator->depth + numerator->height;
        depth = (denominator->height + gap + denominator->depth) - mathAxisOffset;
    }
    float GetCaretX(bool isActive) const override {
        if (!isActive) return width;

        float denX = (width - denominator->width) / 2.0f;
        return denX + denominator->GetCaretX(isActive);
    }
};

class ScriptBox : public Box {
    std::shared_ptr<Box> baseBox;
    std::shared_ptr<Box> superscript;
    std::shared_ptr<Box> subscript;
public:
    ScriptBox(std::shared_ptr<Box> base, std::shared_ptr<Box> sup, std::shared_ptr<Box> sub)
        : baseBox(base), superscript(sup), subscript(sub) {
        RecalculateDimensions();
    }
    void RecalculateDimensions() {
        width = baseBox->width;
        float scriptWidth = 0.0f;
        if (superscript) scriptWidth = std::max(scriptWidth, superscript->width);
        if (subscript)   scriptWidth = std::max(scriptWidth, subscript->width);
        width += scriptWidth + 2.0f;
        height = baseBox->height;
        depth = baseBox->depth;
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
        float currentY = y + shift;
        baseBox->Draw(context, x, currentY);
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
    float GetCaretX(bool isActive) const override {
        if (!isActive) return width;
        float scriptStartX = baseBox->width + 2.0f;
        if (superscript) return scriptStartX + superscript->GetCaretX(isActive);
        if (subscript) return scriptStartX + subscript->GetCaretX(isActive);
        return scriptStartX;
    }
};

class StringBox : public Box {
public:
    std::wstring text;
    float fontSize;
    bool isItalic;
    StringBox(const std::wstring& t, float size, bool italic, IMathRendererContext* ctx)
        : text(t), fontSize(size), isItalic(italic) {
        ctx->MeasureGlyph(text, fontSize, isItalic, width, height, depth);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        context->DrawGlyph(text, x, y + shift, fontSize, isItalic);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        fontSize = newSize;
        ctx->MeasureGlyph(text, fontSize, isItalic, width, height, depth);
    }
};

class RootBox : public Box {
    std::shared_ptr<Box> innerBox;
    float padLeft = 14.0f;
    float padTop = 3.0f;
    float padRight = 2.0f;
    float actualOpeningHeight = 0.0f;
    float actualOpeningDepth = 0.0f;
public:
    RootBox(std::shared_ptr<Box> inner) : innerBox(inner) {
        UpdateMetrics(36.0f);
    }
    void UpdateMetrics(float fontSize) {
        padLeft = fontSize * (14.0f / 36.0f);
        padTop = fontSize * (3.0f / 36.0f);
        padRight = fontSize * (2.0f / 36.0f);
        float minInternalHeight = fontSize * (26.0f / 36.0f);
        float minInternalDepth = fontSize * (8.0f / 36.0f);
        actualOpeningHeight = std::max(innerBox->height, minInternalHeight);
        actualOpeningDepth = std::max(innerBox->depth, minInternalDepth);
        width = padLeft + innerBox->width + padRight;
        height = actualOpeningHeight + padTop;
        depth = actualOpeningDepth + 2.0f;
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift;
        innerBox->Draw(context, x + padLeft, currentY);
        float p1x = x;
        float p1y = currentY - actualOpeningHeight * 0.3f;
        float p2x = x + padLeft * 0.45f;
        float p2y = currentY + actualOpeningDepth + 2.0f;
        float p3x = x + padLeft - 2.0f;
        float p3y = currentY - actualOpeningHeight - padTop + 1.0f;
        float p4x = x + width;
        float p4y = p3y;
        context->DrawLine(p1x, p1y, p2x, p2y, 1.5f);
        context->DrawLine(p2x, p2y, p3x, p3y, 1.5f);
        context->DrawLine(p3x, p3y, p4x, p4y, 1.5f);
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        innerBox->SetFontSize(newSize, ctx);
        UpdateMetrics(newSize);
    }
    float GetCaretX(bool isActive) const override {
        if (!isActive) return width;
        return padLeft + innerBox->GetCaretX(isActive);
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
};

ComPtr<ID2D1Factory7> g_pD2DFactory;
ComPtr<IDWriteFactory3> g_pDWriteFactory;
ComPtr<ID2D1HwndRenderTarget> g_pRenderTarget;
std::shared_ptr<Box> g_pLayoutTree;
std::wstring g_inputText = L"";
int g_cursorPos = 0;
int g_selectionStart = 0;
std::vector<float> g_charOffsets;
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
    if (op == L'@') return 3.0f;
    if (op == L'/') return 2.0f;
    if (op == L'+' || op == L'-') return 1.0f;
    return 0.0f;
}
struct CaretContext {
    float yShift = 0.0f;
    float scale = 1.0f;
    bool isActive = false;
};
CaretContext GetCaretContext(const std::wstring& text, int cursorPos) {
    CaretContext ctx;
    if (cursorPos == 0) return ctx;
    int parenDepth = 0;
    for (int i = cursorPos - 1; i >= 0; i--) {
        wchar_t ch = text[i];
        if (ch == L')') { parenDepth++; continue; }
        if (ch == L'(') {
            parenDepth--;
            if (parenDepth < 0) parenDepth = 0;
            continue;
        }
        if (parenDepth > 0) continue;
        if (ch == L'^') {
            ctx.yShift -= 17.0f * ctx.scale;
            ctx.scale *= 0.65f;
            ctx.isActive = true;
        }
        else if (ch == L'√') {
            ctx.isActive = true;
        }
        else if (ch == L'/') {
            ctx.yShift += 19.0f * ctx.scale;
            ctx.isActive = true;
        }
        else if (ch == L'+' || ch == L'-' || ch == L' ' || ch == L'\x200B' || ch == L'=') {
            break;
        }
    }
    return ctx;
}
bool ApplyOperator(std::stack<wchar_t>& opStack, std::stack<std::shared_ptr<Box>>& boxStack, IMathRendererContext* ctx) {
    if (opStack.empty()) return false;
    wchar_t op = opStack.top();
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
        horiz->Add(std::make_shared<CharBox>(L"!", 36.0f, false, ctx));
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
        horiz->Add(std::make_shared<CharBox>(std::wstring(1, opChar), normalSize, false, ctx));
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
double EvalExpr(const wchar_t*& p);
double EvalFactor(const wchar_t*& p) {
    while (*p == L' ' || *p == L'\x200B') p++;
    if (*p == L'-') { p++; return -EvalFactor(p); }
    double val = 0.0;
    if (*p == L'(') {
        p++; val = EvalExpr(p);
        if (*p == L')') p++;
    }
    else if (wcsncmp(p, L"sin", 3) == 0) { p += 3; val = std::sin(EvalFactor(p)); }
    else if (wcsncmp(p, L"cos", 3) == 0) { p += 3; val = std::cos(EvalFactor(p)); }
    else if (wcsncmp(p, L"log", 3) == 0) { p += 3; val = std::log10(EvalFactor(p)); }
    else if (*p == L'√') { p++; val = std::sqrt(EvalFactor(p)); }
    else if (iswdigit(*p) || *p == L'.') {
        wchar_t* end;
        val = wcstod(p, &end);
        p = end;
    }
    else if (iswalpha(*p)) {
        p++; val = 1.0;
    }
    while (*p == L' ' || *p == L'\x200B') p++;
    while (*p == L'!') {
        p++;
        val = std::tgamma(val + 1.0);
        while (*p == L' ' || *p == L'\x200B') p++;
    }
    if (*p == L'^') {
        p++;
        val = std::pow(val, EvalFactor(p));
    }
    return val;
}
double EvalTerm(const wchar_t*& p) {
    double val = EvalFactor(p);
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
double EvalExpr(const wchar_t*& p) {
    while (*p == L' ' || *p == L'\x200B') p++;
    double val = EvalTerm(p);
    while (true) {
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'+') { p++; val += EvalTerm(p); }
        else if (*p == L'-') { p++; val -= EvalTerm(p); }
        else break;
    }
    return val;
}
std::wstring CalculateResult(std::wstring text) {
    text.erase(std::remove(text.begin(), text.end(), L'\x200B'), text.end());
    if (text.empty()) return L"";
    if (text.find(L'=') != std::wstring::npos) return L"";
    bool hasOperator = false;
    for (wchar_t ch : text) {
        if (ch == L'+' || ch == L'-' || ch == L'*' || ch == L'/' ||
            ch == L'^' || ch == L'√' || ch == L'(' || ch == L')' ||
            ch == L'!' ||
            iswalpha(ch)) {
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
    double result = EvalExpr(p);
    while (*p == L' ' || *p == L'\x200B') p++;
    if (*p != L'\0') return L"";
    if (std::isnan(result) || std::isinf(result)) return L"";
    wchar_t buf[64];
    swprintf_s(buf, 64, L"= %g", result);
    return std::wstring(buf);
}

std::shared_ptr<Box> ParseMathText(const std::wstring& text, IMathRendererContext* ctx) {
    if (text.empty()) return std::make_shared<HorizontalBox>();
    std::stack<std::shared_ptr<Box>> boxStack;
    std::stack<wchar_t> opStack;
    std::stack<size_t> parenBoxStackSizes;
    float normalSize = 36.0f;
    bool expectOperand = true;
    auto EvaluateTop = [&](std::stack<wchar_t>& ops, std::stack<std::shared_ptr<Box>>& boxes) {
        bool pushedDummy = false;
        if (expectOperand) {
            boxes.push(std::make_shared<HorizontalBox>());
            expectOperand = false;
            pushedDummy = true;
        }
        if (!ApplyOperator(ops, boxes, ctx)) {
            if (pushedDummy) boxes.pop();
            wchar_t failedOp = ops.top();
            ops.pop();
            if (failedOp != L'@') {
                boxes.push(std::make_shared<CharBox>(std::wstring(1, failedOp), normalSize, false, ctx));
            }
        }
        };
    for (size_t i = 0; i < text.length(); ) {
        wchar_t ch = text[i];
        if (ch == L' ' || ch == L'\x200B') {
            if (!expectOperand) {
                while (!opStack.empty() && GetIsp(opStack.top()) >= 2.0f) EvaluateTop(opStack, boxStack);
            }
            i++; continue;
        }
        if (iswdigit(ch) || ch == L'.') {
            if (!expectOperand) {
                wchar_t op = L'@';
                while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                opStack.push(op);
            }
            std::wstring numStr = L"";
            while (i < text.length() && (iswdigit(text[i]) || text[i] == L'.')) {
                numStr += text[i]; i++;
            }
            boxStack.push(std::make_shared<CharBox>(numStr, normalSize, false, ctx));
            expectOperand = false;
        }
        else if (iswalpha(ch)) {
            std::wstring funcName;
            if (IsFunctionName(text, i, funcName)) {
                if (!expectOperand) {
                    wchar_t op = L'@';
                    while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                    opStack.push(op);
                }
                boxStack.push(std::make_shared<CharBox>(funcName, normalSize, false, ctx));
                expectOperand = false;
                i += funcName.length();
            }
            else {
                if (!expectOperand) {
                    wchar_t op = L'@';
                    while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                    opStack.push(op);
                }
                boxStack.push(std::make_shared<CharBox>(std::wstring(1, ch), normalSize, true, ctx));
                expectOperand = false;
                i++;
            }
        }
        else if (ch == L'√') {
            if (!expectOperand) {
                wchar_t op = L'@';
                while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                opStack.push(op);
            }
            opStack.push(ch);
            expectOperand = true; i++;
        }
        else if (ch == L'!') {
            if (expectOperand) {
                boxStack.push(std::make_shared<HorizontalBox>());
            }
            else {
                while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(ch)) EvaluateTop(opStack, boxStack);
            }
            opStack.push(ch);
            expectOperand = false;
            i++;
        }
        else if (ch == L'(') {
            if (!expectOperand) {
                wchar_t op = L'@';
                while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                opStack.push(op);
            }
            opStack.push(ch);
            parenBoxStackSizes.push(boxStack.size());
            expectOperand = true; i++;
        }
        else if (ch == L')') {
            while (!opStack.empty() && opStack.top() != L'(') {
                EvaluateTop(opStack, boxStack);
            }

            if (!opStack.empty() && opStack.top() == L'(') {
                opStack.pop();
                size_t baseSize = parenBoxStackSizes.top();
                parenBoxStackSizes.pop();
                auto wrap = std::make_shared<HorizontalBox>();
                wrap->Add(std::make_shared<CharBox>(L"(", normalSize, false, ctx));
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
                wrap->Add(std::make_shared<CharBox>(L")", normalSize, false, ctx));
                boxStack.push(wrap);
            }
            else {
                if (!expectOperand) {
                    wchar_t op = L'@';
                    while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                    opStack.push(op);
                }
                boxStack.push(std::make_shared<CharBox>(L")", normalSize, false, ctx));
            }
            expectOperand = false; i++;
        }
        else if (ch == L'^' || ch == L'/' || ch == L'+' || ch == L'-') {
            while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(ch)) EvaluateTop(opStack, boxStack);
            opStack.push(ch);
            expectOperand = true; i++;
        }
        else {
            if (!expectOperand) {
                wchar_t op = L'@';
                while (!opStack.empty() && GetIsp(opStack.top()) >= GetIcp(op)) EvaluateTop(opStack, boxStack);
                opStack.push(op);
            }
            boxStack.push(std::make_shared<CharBox>(std::wstring(1, ch), normalSize, false, ctx));
            expectOperand = false; i++;
        }
    }
    while (!opStack.empty()) {
        wchar_t op = opStack.top();
        if (op == L'(') {
            opStack.pop();
            size_t baseSize = parenBoxStackSizes.top();
            parenBoxStackSizes.pop();
            auto wrap = std::make_shared<HorizontalBox>();
            wrap->Add(std::make_shared<CharBox>(L"(", normalSize, false, ctx));
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
        EvaluateTop(opStack, boxStack);
    }
    if (boxStack.empty()) return std::make_shared<HorizontalBox>();
    if (boxStack.size() == 1) return boxStack.top();
    auto root = std::make_shared<HorizontalBox>();
    std::vector<std::shared_ptr<Box>> leftover;
    while (!boxStack.empty()) { leftover.push_back(boxStack.top()); boxStack.pop(); }
    for (auto it = leftover.rbegin(); it != leftover.rend(); ++it) root->Add(*it);
    return root;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        SetTimer(hWnd, CARET_TIMER_ID, 500, nullptr);
        break;
    case WM_TIMER:
        if (wParam == CARET_TIMER_ID) {
            g_caretVisible = !g_caretVisible;
            InvalidateRect(hWnd, nullptr, FALSE);
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
                CaretContext ctx = GetCaretContext(g_inputText, g_cursorPos);
                if (ctx.isActive) {
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
                g_inputText.insert(g_cursorPos, L"√"); // Unicodeのルート記号を直接挿入
                g_cursorPos++;
                g_selectionStart = g_cursorPos;
                g_pLayoutTree.reset();
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
        g_pRenderTarget->Clear(D2D1::ColorF(D2D1::ColorF::White));
        Direct2DContext ctx(g_pRenderTarget.Get(), g_pDWriteFactory.Get());
        g_charOffsets.clear();
        for (int i = 0; i <= g_inputText.length(); i++) {
            std::wstring subStr = g_inputText.substr(0, i);
            auto subTree = ParseMathText(subStr, &ctx);
            CaretContext c = GetCaretContext(g_inputText, i);
            g_charOffsets.push_back(subTree->GetCaretX(c.isActive));
        }
        if (!g_pLayoutTree) g_pLayoutTree = ParseMathText(g_inputText, &ctx);
        float startX = 50.0f;
        float startY = 200.0f;
        if (g_cursorPos != g_selectionStart) {
            int startIdx = std::min(g_cursorPos, g_selectionStart);
            int endIdx = std::max(g_cursorPos, g_selectionStart);
            float selStartX = startX + g_charOffsets[startIdx];
            float selEndX = startX + g_charOffsets[endIdx];
            ComPtr<ID2D1SolidColorBrush> highlightBrush;
            g_pRenderTarget->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.4f, 1.0f, 0.2f), &highlightBrush);
            D2D1_RECT_F selRect = D2D1::RectF(selStartX, startY - 36.0f, selEndX, startY + 28.0f);
            g_pRenderTarget->FillRectangle(&selRect, highlightBrush.Get());
        }
        g_pLayoutTree->Draw(&ctx, startX, startY);
        if (!g_resultText.empty()) {
            if (!g_pResultTree) {
                g_pResultTree = ParseMathText(g_resultText, &ctx);
            }
            float resultX = startX + g_pLayoutTree->width + 30.0f;
            g_pResultTree->Draw(&ctx, resultX, startY);
        }
        if (g_caretVisible) {
            float caretX = startX + g_charOffsets[g_cursorPos];
            CaretContext context = GetCaretContext(g_inputText, g_cursorPos);
            float baseTop = -26.0f;
            float baseBottom = 8.0f;
            float caretTop = startY + (baseTop * context.scale) + context.yShift;
            float caretBottom = startY + (baseBottom * context.scale) + context.yShift;
            ctx.DrawLine(caretX + 2.0f, caretTop, caretX + 2.0f, caretBottom, 2.0f);
        }
        g_pRenderTarget->EndDraw();
        EndPaint(hWnd, &ps);
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
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.lpszClassName = L"4mulaWindow";
    RegisterClassExW(&wcex);
    HWND hWnd = CreateWindowW(L"4mulaWindow", L"4mula", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, 800, 450, nullptr, nullptr, hInstance, nullptr);
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