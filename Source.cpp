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
#include <fstream>
#include <shlobj.h>
#include <boost/multiprecision/cpp_dec_float.hpp>
#include <boost/math/special_functions/gamma.hpp>
#include <boost/multiprecision/cpp_int.hpp>
#include <boost/multiprecision/miller_rabin.hpp>
#include <boost/random.hpp>
#include <map>
#include <ctime>
#include "resource.h"

using namespace boost::multiprecision;
typedef boost::multiprecision::number<boost::multiprecision::cpp_dec_float<500>> BigFloat;
using namespace Microsoft::WRL;

class Box;
std::wstring g_resultText = L"";
std::shared_ptr<Box> g_pResultTree = nullptr;
bool g_isDragging = false;
bool g_isDarkMode = false;

// -----------------------------------------------------
// 追加: 保存リストとUIステート管理
// -----------------------------------------------------
struct SavedFormula {
    std::wstring formula;
    std::wstring result;
    std::shared_ptr<Box> pFormulaTree;
    std::shared_ptr<Box> pResultTree;
};
std::vector<SavedFormula> g_savedFormulas;
int g_selectedIndex = -1; // -1 は何も選択されていない状態
float g_scrollOffsetY = 0.0f;
float g_maxScrollY = 0.0f;
bool g_isDraggingScrollbar = false;
float g_dragStartY = 0.0f;
float g_dragStartScrollY = 0.0f;
D2D1_RECT_F g_pinButtonRect = { 0, 0, 0, 0 };
D2D1_RECT_F g_scrollbarThumbRect = { 0, 0, 0, 0 };

const float INPUT_AREA_HEIGHT = 120.0f;
const float LIST_ITEM_HEIGHT = 80.0f;

std::string WideToUTF8(const std::wstring& wstr) {
    if (wstr.empty()) return "";
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

std::wstring UTF8ToWide(const std::string& str) {
    if (str.empty()) return L"";
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

std::wstring GetSaveFilePath() {
    PWSTR path = nullptr;
    std::wstring result = L"4mula_saved.txt"; // 取得失敗時のフォールバック
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &path))) {
        std::wstring dir = std::wstring(path) + L"\\4mula";
        CreateDirectoryW(dir.c_str(), nullptr);
        result = dir + L"\\4mula_saved.txt";
        CoTaskMemFree(path);
    }
    return result;
}

void SaveFormulasToFile() {
    std::wstring path = GetSaveFilePath();
    std::ofstream out(path, std::ios::binary);
    if (!out) return;

    for (const auto& item : g_savedFormulas) {
        out << WideToUTF8(item.formula) << "\t" << WideToUTF8(item.result) << "\n";
    }
}

void LoadFormulasFromFile() {
    std::wstring path = GetSaveFilePath();
    std::ifstream in(path, std::ios::binary);
    if (!in) return;

    std::string line;
    while (std::getline(in, line)) {
        std::wstring wline = UTF8ToWide(line);
        size_t tabPos = wline.find(L'\t');
        if (tabPos != std::wstring::npos) {
            SavedFormula sf;
            sf.formula = wline.substr(0, tabPos);
            sf.result = wline.substr(tabPos + 1);
            g_savedFormulas.push_back(sf);
        }
    }
}

class IMathRendererContext {
public:
    virtual ~IMathRendererContext() = default;
    virtual void DrawGlyph(const std::wstring& text, float x, float y, float fontSize, bool isItalic) = 0;
    virtual void DrawLine(float x1, float y1, float x2, float y2, float thickness) = 0;
    virtual void FillCircle(float x, float y, float r) = 0;
    virtual void FillPolygon(const std::vector<std::pair<float, float>>& vertices) = 0;
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
extern int g_cursorPos;

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
public:
    std::vector<std::shared_ptr<Box>> children;
    bool isParenGroup = false;
    bool hideParens = false;
    bool ShouldHide() const {
        if (!hideParens) return false;
        if (g_cursorPos >= srcStart && g_cursorPos <= srcEnd) {
            return false;
        }
        return true;
    }
    void Add(std::shared_ptr<Box> box) {
        children.push_back(box);
        width += box->width;
        height = std::max(height, box->height - box->shift);
        depth = std::max(depth, box->depth + box->shift);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentX = x; float currentY = y + shift;
        bool doHide = ShouldHide();
        for (size_t i = 0; i < children.size(); ++i) {
            if (doHide && (i == 0 || i == children.size() - 1)) {
            }
            else {
                children[i]->Draw(context, currentX, currentY);
            }
            currentX += children[i]->width;
        }
    }
    void SetFontSize(float newSize, IMathRendererContext* ctx) override {
        width = 0.0f; height = 0.0f; depth = 0.0f;
        bool doHide = ShouldHide();
        for (size_t i = 0; i < children.size(); ++i) {
            children[i]->SetFontSize(newSize, ctx);
            if (doHide && (i == 0 || i == children.size() - 1)) {
                children[i]->width = 0.0f;
            }
            width += children[i]->width;

            if (!(doHide && (i == 0 || i == children.size() - 1))) {
                height = std::max(height, children[i]->height - children[i]->shift);
                depth = std::max(depth, children[i]->depth + children[i]->shift);
            }
        }
    }
    void MapCaret(IMathRendererContext* ctx, float x, float y, float scale, bool isActive, std::vector<CaretMetrics>& metrics) const override {
        if (children.empty() && srcStart >= 0 && srcStart < metrics.size()) {
            if (metrics[srcStart].x < 0) {
                metrics[srcStart].x = x; metrics[srcStart].y = y + shift;
                metrics[srcStart].scale = scale; metrics[srcStart].isActive = isActive;
            }
            if (srcEnd >= 0 && srcEnd < metrics.size() && metrics[srcEnd].x < 0) {
                metrics[srcEnd].x = x; metrics[srcEnd].y = y + shift;
                metrics[srcEnd].scale = scale; metrics[srcEnd].isActive = isActive;
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
    float fontSize = 36.0f;
    float padLeft = 14.0f; float padTop = 3.0f; float padRight = 2.0f;
    float actualOpeningHeight = 0.0f; float actualOpeningDepth = 0.0f;
public:
    RootBox(std::shared_ptr<Box> inner) : innerBox(inner) { UpdateMetrics(36.0f); }
    void UpdateMetrics(float newFontSize) {
        fontSize = newFontSize;
        padLeft = fontSize * (16.0f / 36.0f);
        padTop = fontSize * (3.0f / 36.0f); padRight = fontSize * (2.0f / 36.0f);
        float minInternalHeight = fontSize * (26.0f / 36.0f); float minInternalDepth = fontSize * (8.0f / 36.0f);
        actualOpeningHeight = std::max(innerBox->height, minInternalHeight); actualOpeningDepth = std::max(innerBox->depth, minInternalDepth);
        width = padLeft + innerBox->width + padRight; height = actualOpeningHeight + padTop;
        depth = actualOpeningDepth + fontSize * (2.0f / 36.0f);
    }
    void Draw(IMathRendererContext* context, float x, float y) const override {
        float currentY = y + shift; innerBox->Draw(context, x + padLeft, currentY);
        float thicknessOver = fontSize * (1.2f / 36.0f);
        float thicknessTickUp = fontSize * (1.2f / 36.0f);
        float thicknessTickDown = fontSize * (2.8f / 36.0f);
        float thicknessMain = fontSize * (1.8f / 36.0f);
        float topY = currentY - actualOpeningHeight - padTop;
        float bottomY = currentY + actualOpeningDepth;
        float p0x = x + padLeft * 0.05f;
        float p0y = currentY - actualOpeningHeight * 0.30f;
        float p1x = x + padLeft * 0.35f;
        float p1y = currentY - actualOpeningHeight * 0.38f;
        float p2x = x + padLeft * 0.60f;
        float p2y = bottomY + fontSize * (3.0f / 36.0f);
        float overLineY = topY + thicknessOver * 0.5f;
        float p3x = x + padLeft - thicknessMain * 0.5f;
        float p3y = overLineY - thicknessMain * 0.2f;
        float p4x = x + width;
        context->DrawLine(p0x, p0y, p1x, p1y, thicknessTickUp);
        context->DrawLine(p1x, p1y, p2x, p2y, thicknessTickDown);
        context->DrawLine(p2x, p2y, p3x, p3y, thicknessMain);
        float overLineStartX = p3x - thicknessMain * 0.5f + thicknessOver * 0.5f;
        context->DrawLine(overLineStartX, overLineY, p4x, overLineY, thicknessOver);
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
        m_pDWriteFactory->CreateTextFormat(L"Cambria", nullptr, DWRITE_FONT_WEIGHT_NORMAL,
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
    void FillPolygon(const std::vector<std::pair<float, float>>& vertices) override {
        if (vertices.size() < 3) return;
        ComPtr<ID2D1Factory> pFactory;
        m_pRT->GetFactory(&pFactory);
        ComPtr<ID2D1GeometrySink> pSink;
        ComPtr<ID2D1PathGeometry> pPathGeometry;
        pFactory->CreatePathGeometry(&pPathGeometry);
        pPathGeometry->Open(&pSink);
        pSink->BeginFigure(D2D1::Point2F(vertices[0].first, vertices[0].second), D2D1_FIGURE_BEGIN_FILLED);
        for (size_t i = 1; i < vertices.size(); ++i) {
            pSink->AddLine(D2D1::Point2F(vertices[i].first, vertices[i].second));
        }
        pSink->EndFigure(D2D1_FIGURE_END_CLOSED);
        pSink->Close();
        m_pRT->FillGeometry(pPathGeometry.Get(), m_pBrush.Get());
    }
    void SetTextColor(float r, float g, float b, float a = 1.0f) override {
        m_pBrush->SetColor(D2D1::ColorF(r, g, b, a));
    }
};

ComPtr<ID2D1Factory7> g_pD2DFactory;
ComPtr<IDWriteFactory3> g_pDWriteFactory;
ComPtr<ID2D1HwndRenderTarget> g_pRenderTarget;
ComPtr<IDWriteTextFormat> g_pUiTextFormat; // 追加
std::shared_ptr<Box> g_pLayoutTree;
std::wstring g_inputText = L"";
int g_cursorPos = 0;
int g_selectionStart = 0;
bool g_caretVisible = true;
const UINT CARET_TIMER_ID = 1;
const UINT DEMO_TYPING_TIMER_ID = 2;
const UINT DEMO_PAUSE_TIMER_ID = 3;
bool g_isDemoMode = false;
size_t g_demoIndex = 0;
size_t g_demoCharIndex = 0;
std::vector<std::wstring> g_demoFormulas;

void InitDemoFormulas() {
    g_demoFormulas = {
        L"version",
        L"1+1",
        L"(10*(10+1))/2",
        L"987654321/123456789",
        L"((1+√(5))/2)^10/√(5)",
        L"2305843009213693951",
        L"tan(1)-(sin(1)/cos(1))",
        L"ln(exp(10))",
        L"mod(2026^365,7)",
        L"√(12345678987654321)",
        L"√(2+√(2+√(2+√(2))))",
        L"(1+1/10000)^10000",
        L"2*3.14159*√(2/9.8)",
        L"4/3*3.14159*6371^3",
        L"1+1/(1+1/(1+1/(1+1)))",
        L"1/(√(2*3.14159))*exp(-(0^2)/2)",
        L"6.674e-11*5.972e24/(6.371e6^2)",
        L"　　4mula"
    };
}

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
    if (op == L'!' || op == L'%') return 4.8f;
    if (op == L'√') return 4.5f;
    if (op == L'^') return 4.0f;
    if (op == L'_' || op == L'#') return 3.5f;
    if (op == L'@') return 3.0f;
    if (op == L'/') return 2.0f;
    if (op == L'*') return 1.5f;
    if (op == L'+' || op == L'-') return 1.0f;
    return 0.0f;
}

float GetIsp(wchar_t op) {
    if (op == L'(') return 0.0f;
    if (op == L'!' || op == L'%') return 4.8f;
    if (op == L'√') return 1.5f;
    if (op == L'^') return 1.5f;
    if (op == L'_' || op == L'#') return 3.5f;
    if (op == L'@') return 3.0f;
    if (op == L'/') return 1.5f;
    if (op == L'*') return 1.5f;
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
    auto stripParens = [&](std::shared_ptr<Box> b) {
        if (auto h = std::dynamic_pointer_cast<HorizontalBox>(b)) {
            if (h->isParenGroup) {
                h->hideParens = true;
                h->SetFontSize(36.0f, ctx);
            }
        }
        };
    if (op == L'√') {
        if (boxStack.empty()) return false;
        opStack.pop();
        std::shared_ptr<Box> inner = boxStack.top(); boxStack.pop();
        stripParens(inner);
        boxStack.push(std::make_shared<RootBox>(inner));
        return true;
    }
    if (op == L'!' || op == L'%') {
        if (boxStack.empty()) return false;
        opStack.pop();
        std::shared_ptr<Box> inner = boxStack.top(); boxStack.pop();
        auto horiz = std::make_shared<HorizontalBox>();
        horiz->Add(inner);
        auto postfixBox = std::make_shared<CharBox>(std::wstring(1, op), 36.0f, false, ctx);
        postfixBox->srcStart = opInfo.idx;
        postfixBox->srcEnd = opInfo.idx + 1;
        horiz->Add(postfixBox);
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
        stripParens(right);
        right->SetFontSize(scriptSize, ctx);
        result = std::make_shared<ScriptBox>(left, right, nullptr);
        break;
    case L'*':
    {
        auto horiz = std::make_shared<HorizontalBox>();
        horiz->Add(left);
        horiz->Add(std::make_shared<SpaceBox>(4.0f));
        auto opBox = std::make_shared<CharBox>(std::wstring(1, L'\x22C5'), normalSize, false, ctx);
        opBox->srcStart = opInfo.idx;
        opBox->srcEnd = opInfo.idx + 1;
        horiz->Add(opBox);
        horiz->Add(std::make_shared<SpaceBox>(4.0f));
        horiz->Add(right);
        result = horiz;
        break;
    }
    case L'/':
        stripParens(left);
        stripParens(right);
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
        L"log10", L"log2", L"powmod",
        L"sin", L"cos", L"tan", L"log", L"lim", L"exp", L"max", L"min", L"det", L"mod",
        L"abs", L"ceil", L"floor", L"round",
        L"ln", L"is", L"prime"
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
BigFloat EvalTerm(const wchar_t*& p);
BigFloat EvalGroup(const wchar_t*& p);
BigFloat EvalPrimary(const wchar_t*& p);

BigFloat EvalPrimary(const wchar_t*& p) {
    while (*p == L' ' || *p == L'\x200B') p++;
    BigFloat val = 0.0;
    if (*p == L'(') {
        p++; val = EvalExpr(p);
        if (*p == L')') p++;
    }
    else if (wcsncmp(p, L"arcsin", 6) == 0) { p += 6; val = asin(EvalPrimary(p)); }
    else if (wcsncmp(p, L"arccos", 6) == 0) { p += 6; val = acos(EvalPrimary(p)); }
    else if (wcsncmp(p, L"arctan", 6) == 0) { p += 6; val = atan(EvalPrimary(p)); }
    else if (wcsncmp(p, L"sinh", 4) == 0) { p += 4; val = sinh(EvalPrimary(p)); }
    else if (wcsncmp(p, L"cosh", 4) == 0) { p += 4; val = cosh(EvalPrimary(p)); }
    else if (wcsncmp(p, L"tanh", 4) == 0) { p += 4; val = tanh(EvalPrimary(p)); }
    else if (wcsncmp(p, L"asin", 4) == 0) { p += 4; val = asin(EvalPrimary(p)); }
    else if (wcsncmp(p, L"acos", 4) == 0) { p += 4; val = acos(EvalPrimary(p)); }
    else if (wcsncmp(p, L"atan", 4) == 0) { p += 4; val = atan(EvalPrimary(p)); }
    else if (wcsncmp(p, L"sin", 3) == 0) { p += 3; val = sin(EvalPrimary(p)); }
    else if (wcsncmp(p, L"cos", 3) == 0) { p += 3; val = cos(EvalPrimary(p)); }
    else if (wcsncmp(p, L"tan", 3) == 0) { p += 3; val = tan(EvalPrimary(p)); }
    else if (wcsncmp(p, L"log10", 5) == 0) { p += 5; val = log10(EvalPrimary(p)); }
    else if (wcsncmp(p, L"log2", 4) == 0) { p += 4; val = log2(EvalPrimary(p)); }
    else if (wcsncmp(p, L"log", 3) == 0) { p += 3; val = log10(EvalPrimary(p)); }
    else if (wcsncmp(p, L"exp", 3) == 0) { p += 3; val = exp(EvalPrimary(p)); }
    else if (wcsncmp(p, L"ln", 2) == 0) { p += 2; val = log(EvalPrimary(p)); }
    else if (wcsncmp(p, L"lim", 3) == 0) {
        p += 3;
        val = EvalPrimary(p);
    }
    else if (wcsncmp(p, L"max", 3) == 0 || wcsncmp(p, L"min", 3) == 0) {
        std::wstring fName(p, 3);
        p += 3;
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'(') {
            p++;
            val = EvalExpr(p);
            while (true) {
                while (*p == L' ' || *p == L'\x200B') p++;
                if (*p == L',') {
                    p++;
                    BigFloat nextVal = EvalExpr(p);
                    if (fName == L"max") { if (nextVal > val) val = nextVal; }
                    else if (fName == L"min") { if (nextVal < val) val = nextVal; }
                }
                else {
                    break;
                }
            }
            if (*p == L')') p++;
        }
        else {
            val = std::numeric_limits<BigFloat>::quiet_NaN();
        }
    }
    else if (wcsncmp(p, L"mod", 3) == 0) {
        p += 3;
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'(') {
            p++;
            const wchar_t* arg1_start = p;
            int p_count = 0;
            const wchar_t* comma_pos = nullptr;
            const wchar_t* scan = p;
            while (*scan != L'\0') {
                if (*scan == L'(') p_count++;
                else if (*scan == L')') {
                    if (p_count == 0) break;
                    p_count--;
                }
                else if (p_count == 0 && *scan == L',') {
                    comma_pos = scan;
                    break;
                }
                scan++;
            }
            if (comma_pos != nullptr) {
                const wchar_t* temp_p = arg1_start;
                while (*temp_p == L' ' || *temp_p == L'\x200B') temp_p++;
                bool is_negative = false;
                if (*temp_p == L'-') { is_negative = true; temp_p++; }
                else if (*temp_p == L'+') { temp_p++; }
                BigFloat val_base = EvalPrimary(temp_p);
                while (*temp_p == L'!' || *temp_p == L'%') {
                    if (*temp_p == L'!') { temp_p++; val_base = boost::math::tgamma(val_base + 1); }
                    else if (*temp_p == L'%') { temp_p++; val_base /= 100.0; }
                }
                if (is_negative) val_base = -val_base;
                while (*temp_p == L' ' || *temp_p == L'\x200B') temp_p++;
                if (*temp_p == L'^') {
                    temp_p++;
                    BigFloat val_exp = EvalGroup(temp_p);
                    while (*temp_p == L' ' || *temp_p == L'\x200B') temp_p++;
                    if (temp_p == comma_pos) {
                        p = comma_pos + 1;
                        BigFloat val_mod = EvalExpr(p);
                        while (*p == L' ' || *p == L'\x200B') p++;
                        if (*p == L')') p++;
                        if (floor(val_base) == val_base && floor(val_exp) == val_exp && floor(val_mod) == val_mod && val_mod != 0 && val_exp >= 0) {
                            cpp_int base_int = val_base.convert_to<cpp_int>();
                            cpp_int exp_int = val_exp.convert_to<cpp_int>();
                            cpp_int mod_int = val_mod.convert_to<cpp_int>();
                            base_int = base_int % mod_int;
                            if (base_int < 0) base_int += (mod_int > 0 ? mod_int : -mod_int);
                            val = BigFloat(boost::multiprecision::powm(base_int, exp_int, mod_int));
                        }
                        else {
                            val = fmod(pow(val_base, val_exp), val_mod);
                        }
                        return val;
                    }
                }
            }
            BigFloat val1 = EvalExpr(p);
            while (*p == L' ' || *p == L'\x200B') p++;
            if (*p == L',') {
                p++;
                BigFloat val2 = EvalExpr(p);
                while (*p == L' ' || *p == L'\x200B') p++;
                if (*p == L')') p++;
                val = fmod(val1, val2);
            }
            else {
                val = std::numeric_limits<BigFloat>::quiet_NaN();
            }
        }
        else {
            val = std::numeric_limits<BigFloat>::quiet_NaN();
        }
    }
    else if (wcsncmp(p, L"powmod", 6) == 0) {
        p += 6;
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'(') {
            p++;
            BigFloat val1 = EvalExpr(p);
            while (*p == L' ' || *p == L'\x200B') p++;
            if (*p == L',') {
                p++;
                BigFloat val2 = EvalExpr(p);
                while (*p == L' ' || *p == L'\x200B') p++;
                if (*p == L',') {
                    p++;
                    BigFloat val3 = EvalExpr(p);
                    while (*p == L' ' || *p == L'\x200B') p++;
                    if (*p == L')') p++;
                    if (floor(val1) == val1 && floor(val2) == val2 && floor(val3) == val3 && val3 != 0 && val2 >= 0) {
                        cpp_int base = val1.convert_to<cpp_int>();
                        cpp_int exp = val2.convert_to<cpp_int>();
                        cpp_int modulo = val3.convert_to<cpp_int>();
                        cpp_int res = boost::multiprecision::powm(base, exp, modulo);
                        val = BigFloat(res);
                    }
                    else {
                        val = std::numeric_limits<BigFloat>::quiet_NaN();
                    }
                }
                else {
                    val = std::numeric_limits<BigFloat>::quiet_NaN();
                }
            }
            else {
                val = std::numeric_limits<BigFloat>::quiet_NaN();
            }
        }
        else {
            val = std::numeric_limits<BigFloat>::quiet_NaN();
        }
    }
    else if (wcsncmp(p, L"abs", 3) == 0) { p += 3; val = abs(EvalPrimary(p)); }
    else if (wcsncmp(p, L"ceil", 4) == 0) { p += 4; val = ceil(EvalPrimary(p)); }
    else if (wcsncmp(p, L"floor", 5) == 0) { p += 5; val = floor(EvalPrimary(p)); }
    else if (wcsncmp(p, L"round", 5) == 0) {
        p += 5;
        while (*p == L' ' || *p == L'\x200B') p++;
        if (*p == L'(') {
            p++;
            BigFloat val1 = EvalExpr(p);
            while (*p == L' ' || *p == L'\x200B') p++;
            if (*p == L',') {
                p++;
                BigFloat val2 = EvalExpr(p);
                while (*p == L' ' || *p == L'\x200B') p++;
                if (*p == L')') p++;
                BigFloat factor = pow(BigFloat(10), val2);
                val = round(val1 * factor) / factor;
            }
            else {
                if (*p == L')') p++;
                val = round(val1);
            }
        }
        else {
            val = std::numeric_limits<BigFloat>::quiet_NaN();
        }
    }
    else if (*p == L'√') { p++; val = sqrt(EvalGroup(p)); }
    else if (iswdigit(*p) || *p == L'.') {
        if (*p == L'0' && (*(p + 1) == L'x' || *(p + 1) == L'X')) {
            std::string hexStr = "0x";
            p += 2;
            while (iswdigit(*p) || (*p >= L'a' && *p <= L'f') || (*p >= L'A' && *p <= L'F')) {
                hexStr += static_cast<char>(*p);
                p++;
            }
            if (hexStr.length() == 2) {
                val = std::numeric_limits<BigFloat>::quiet_NaN();
            }
            else {
                val = BigFloat(cpp_int(hexStr));
            }
        }
        else {
            std::string numStr;
            while (iswdigit(*p) || *p == L'.') {
                numStr += static_cast<char>(*p);
                p++;
            }
            if (*p == L'e' || *p == L'E') {
                const wchar_t* temp = p + 1;
                if (*temp == L'+' || *temp == L'-') temp++;
                if (iswdigit(*temp)) {
                    numStr += static_cast<char>(*p++);
                    if (*p == L'+' || *p == L'-') {
                        numStr += static_cast<char>(*p++);
                    }
                    while (iswdigit(*p)) {
                        numStr += static_cast<char>(*p++);
                    }
                }
            }
            val = BigFloat(numStr);
        }
    }
    else if (iswalpha(*p)) {
        p++;
        val = std::numeric_limits<BigFloat>::quiet_NaN();
    }
    return val;
}

BigFloat EvalFactor(const wchar_t*& p) {
    while (*p == L' ' || *p == L'\x200B') p++;
    if (*p == L'-') { p++; return -EvalFactor(p); }
    if (*p == L'+') { p++; return EvalFactor(p); }
    BigFloat val = EvalPrimary(p);
    while (*p == L'!' || *p == L'%') {
        if (*p == L'!') {
            p++;
            if (val > 1000.0 || val < 0.0) {
                return std::numeric_limits<BigFloat>::quiet_NaN();
            }
            val = boost::math::tgamma(val + 1);
        }
        else if (*p == L'%') {
            p++;
            val /= 100.0;
        }
    }
    if (*p == L'^') {
        p++;
        BigFloat exponent = EvalGroup(p);
        if (abs(exponent) > 10000.0) {
            val = std::numeric_limits<BigFloat>::quiet_NaN();
        }
        else {
            val = pow(val, exponent);
        }
    }
    return val;
}

BigFloat EvalGroup(const wchar_t*& p) {
    BigFloat val = EvalFactor(p);
    while (true) {
        if (*p == L' ' || *p == L'\x200B') break;
        if (*p == L'*') { p++; val *= EvalFactor(p); }
        else if (*p == L'/') { p++; val /= EvalFactor(p); }
        else if (iswdigit(*p) || iswalpha(*p) || *p == L'(' || *p == L'√' || wcsncmp(p, L"sin", 3) == 0 || wcsncmp(p, L"cos", 3) == 0) {
            val *= EvalFactor(p);
        }
        else break;
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

using boost::multiprecision::cpp_int;

cpp_int gcd_cpp_int(cpp_int a, cpp_int b) {
    while (b != 0) {
        cpp_int temp = b;
        b = a % b;
        a = temp;
    }
    return a;
}

cpp_int pollard_rho(cpp_int n) {
    if (n % 2 == 0) return 2;
    cpp_int x = 2, y = 2, d = 1, c = 1;
    auto f = [&n, &c](const cpp_int& val) { return (val * val + c) % n; };
    while (d == 1) {
        x = f(x);
        y = f(f(y));
        cpp_int diff = x > y ? x - y : y - x;
        d = gcd_cpp_int(diff, n);
        if (d == n) {
            c++;
            x = 2;
            y = 2;
            d = 1;
        }
    }
    return d;
}

void factorize_cpp_int(cpp_int n, std::map<cpp_int, int>& factors, boost::random::mt19937& rng) {
    if (n <= 1) return;
    if (boost::multiprecision::miller_rabin_test(n, 25, rng)) {
        factors[n]++;
        return;
    }
    cpp_int divisor = pollard_rho(n);
    factorize_cpp_int(divisor, factors, rng);
    factorize_cpp_int(n / divisor, factors, rng);
}

std::wstring CalculateResult(std::wstring text) {
    std::replace(text.begin(), text.end(), L'\x200B', L' ');
    std::wstring cleanText = text;
    cleanText.erase(std::remove(cleanText.begin(), cleanText.end(), L' '), cleanText.end());
    if (cleanText == L"version") {
        return L"v1.0.5";
    }
    if (text.empty()) return L"";
    if (text.find(L'=') != std::wstring::npos) return L"";
    bool hasOperator = false;
    for (wchar_t ch : text) {
        if (ch == L'+' || ch == L'-' || ch == L'*' || ch == L'/' ||
            ch == L'^' || ch == L'√' || ch == L'(' || ch == L')' ||
            ch == L'!' || ch == L'%' || iswalpha(ch)) {
            hasOperator = true;
            break;
        }
    }
    bool isOnlyDigits = !cleanText.empty();
    for (wchar_t ch : cleanText) {
        if (!iswdigit(ch)) {
            isOnlyDigits = false;
            break;
        }
    }
    if (!hasOperator && !isOnlyDigits) return L"";
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
    if (isOnlyDigits && result >= 2 && floor(result) == result) {
        cpp_int n = result.convert_to<cpp_int>();
        boost::random::mt19937 rng(static_cast<uint32_t>(time(0)));
        if (boost::multiprecision::miller_rabin_test(n, 25, rng)) {
            return L"is\x00A0prime";
        }
        else {
            std::map<cpp_int, int> factors;
            factorize_cpp_int(n, factors, rng);
            std::wstring factStr = L"";
            bool first = true;
            for (auto const& pair : factors) {
                if (!first) factStr += L" * ";
                first = false;
                std::string s = pair.first.str();
                factStr += std::wstring(s.begin(), s.end());
                if (pair.second > 1) {
                    factStr += L"^" + std::to_wstring(pair.second);
                }
            }
            return factStr;
        }
    }
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
                while (!opStack.empty() && GetIsp(opStack.top().ch) >= 1.5f) EvaluateTop(opStack, boxStack, static_cast<int>(i));
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
            while (!opStack.empty() && GetIsp(opStack.top().ch) >= 1.5f) EvaluateTop(opStack, boxStack, static_cast<int>(i));
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
        else if (ch == L'!' || ch == L'%') {
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
                wrap->isParenGroup = true;
                wrap->srcStart = leftParenIdx;
                wrap->srcEnd = (int)(i + 1);
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
        else if (ch == L'^' || ch == L'/' || ch == L'*') {
            while (!opStack.empty() && GetIsp(opStack.top().ch) >= GetIcp(ch)) EvaluateTop(opStack, boxStack, static_cast<int>(i));
            opStack.push({ ch, (int)i });
            expectOperand = true; i++;
        }
        else if (ch == L',') {
            while (!opStack.empty() && GetIsp(opStack.top().ch) >= 1.0f) {
                EvaluateTop(opStack, boxStack, static_cast<int>(i));
            }
            auto box = std::make_shared<CharBox>(L",", normalSize, false, ctx);
            box->srcStart = (int)i; box->srcEnd = (int)(i + 1);
            boxStack.push(box);
            expectOperand = true;
            i++;
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
            wrap->isParenGroup = true;
            wrap->srcStart = opInfo.idx;
            wrap->srcEnd = static_cast<int>(text.length());
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
    float startX = 50.0f;
    float startY = INPUT_AREA_HEIGHT / 2.0f;
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
        LoadFormulasFromFile();
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
        else if (wParam == DEMO_TYPING_TIMER_ID && g_isDemoMode) {
            if (g_demoCharIndex < g_demoFormulas[g_demoIndex].length()) {
                g_inputText += g_demoFormulas[g_demoIndex][g_demoCharIndex];
                g_demoCharIndex++;
                g_cursorPos = static_cast<int>(g_inputText.length());
                g_selectionStart = g_cursorPos;
                g_resultText = CalculateResult(g_inputText);
                g_pResultTree.reset();
                g_pLayoutTree.reset();
                g_caretVisible = true;
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            else {
                if (!g_caretMetrics.empty() &&
                    g_cursorPos < g_caretMetrics.size() &&
                    g_caretMetrics[g_cursorPos].isActive) {
                    g_inputText.insert(g_cursorPos, 1, L'\x200B');
                    g_cursorPos++;
                    g_selectionStart = g_cursorPos;
                    g_pLayoutTree.reset();
                    g_caretVisible = true;
                    InvalidateRect(hWnd, nullptr, FALSE);
                }
                else {
                    KillTimer(hWnd, DEMO_TYPING_TIMER_ID);
                    SetTimer(hWnd, DEMO_PAUSE_TIMER_ID, 400, nullptr);
                }
            }
        }
        else if (wParam == DEMO_PAUSE_TIMER_ID && g_isDemoMode) {
            KillTimer(hWnd, DEMO_PAUSE_TIMER_ID);
            if (g_demoIndex >= g_demoFormulas.size() - 1) {
                g_isDemoMode = false;
            }
            else {
                g_demoIndex++;
                g_demoCharIndex = 0;
                g_inputText.clear();
                g_cursorPos = 0;
                g_selectionStart = 0;
                g_resultText = L"";
                g_pResultTree.reset();
                g_pLayoutTree.reset();
                InvalidateRect(hWnd, nullptr, FALSE);
                SetTimer(hWnd, DEMO_TYPING_TIMER_ID, 50, nullptr);
            }
        }
        break;
    case WM_LBUTTONDOWN:
        if (g_isDemoMode) {
            g_inputText.clear();
            g_cursorPos = 0;
            g_selectionStart = 0;
            SetTimer(hWnd, DEMO_TYPING_TIMER_ID, 100, nullptr);
        }
        {
            int xPos = GET_X_LPARAM(lParam);
            int yPos = GET_Y_LPARAM(lParam);

            // ピンボタンの判定
            if (xPos >= g_pinButtonRect.left && xPos <= g_pinButtonRect.right &&
                yPos >= g_pinButtonRect.top && yPos <= g_pinButtonRect.bottom) {
                bool isDup = false;
                for (auto& f : g_savedFormulas) {
                    if (f.formula == g_inputText) { isDup = true; break; }
                }
                if (!isDup && !g_inputText.empty() && !g_resultText.empty()) {
                    SavedFormula sf;
                    sf.formula = g_inputText;
                    sf.result = g_resultText;
                    g_savedFormulas.push_back(sf);
                    SaveFormulasToFile();

                    RECT clientRect;
                    GetClientRect(hWnd, &clientRect);
                    float viewHeight = clientRect.bottom - INPUT_AREA_HEIGHT;
                    float contentHeight = g_savedFormulas.size() * LIST_ITEM_HEIGHT;
                    g_maxScrollY = std::max(0.0f, contentHeight - viewHeight);
                    g_scrollOffsetY = g_maxScrollY; // 一番下へスクロール

                    InvalidateRect(hWnd, nullptr, FALSE);
                }
                return 0;
            }

            // スクロールバーの判定
            if (xPos >= g_scrollbarThumbRect.left && xPos <= g_scrollbarThumbRect.right &&
                yPos >= g_scrollbarThumbRect.top && yPos <= g_scrollbarThumbRect.bottom) {
                g_isDraggingScrollbar = true;
                g_dragStartY = (float)yPos;
                g_dragStartScrollY = g_scrollOffsetY;
                SetCapture(hWnd);
                return 0;
            }

            // リスト領域のクリック判定（ここを変更）
            if (yPos > INPUT_AREA_HEIGHT) {
                RECT rc;
                GetClientRect(hWnd, &rc);

                // クリックされたY座標から、スクロールオフセットを考慮してインデックスを計算
                float listClickY = yPos - INPUT_AREA_HEIGHT + g_scrollOffsetY;
                int itemIndex = static_cast<int>(listClickY / LIST_ITEM_HEIGHT);

                if (itemIndex >= 0 && itemIndex < g_savedFormulas.size()) {
                    g_selectedIndex = itemIndex; // 選択インデックスを更新
                    // 削除ボタンの領域計算
                    float listY = INPUT_AREA_HEIGHT - g_scrollOffsetY + (itemIndex * LIST_ITEM_HEIGHT) + (LIST_ITEM_HEIGHT / 2.0f);
                    float delBtnWidth = 24.0f;
                    float delBtnHeight = 24.0f;
                    float delBtnLeft = rc.right - 60.0f; // スクロールバーと被らない位置
                    float delBtnRight = delBtnLeft + delBtnWidth;
                    float delBtnTop = listY - delBtnHeight / 2.0f;
                    float delBtnBottom = listY + delBtnHeight / 2.0f;

                    // 削除ボタンがクリックされたか判定
                    if (xPos >= delBtnLeft && xPos <= delBtnRight && yPos >= delBtnTop && yPos <= delBtnBottom) {
                        // リストから削除してファイルに保存
                        g_savedFormulas.erase(g_savedFormulas.begin() + itemIndex);
                        SaveFormulasToFile();

                        // アイテムが減ったのでスクロールの最大値を再計算
                        float viewHeight = rc.bottom - INPUT_AREA_HEIGHT;
                        float contentHeight = g_savedFormulas.size() * LIST_ITEM_HEIGHT;
                        g_maxScrollY = std::max(0.0f, contentHeight - viewHeight);
                        if (g_scrollOffsetY > g_maxScrollY) g_scrollOffsetY = g_maxScrollY;

                        InvalidateRect(hWnd, nullptr, FALSE);
                        return 0;
                    }

                    // 削除ボタン以外がクリックされた場合は数式と結果を復元
                    g_inputText = g_savedFormulas[itemIndex].formula;
                    g_resultText = g_savedFormulas[itemIndex].result;

                    g_cursorPos = static_cast<int>(g_inputText.length());
                    g_selectionStart = g_cursorPos;

                    g_pLayoutTree.reset();
                    g_pResultTree.reset();
                    g_caretVisible = true;
                    InvalidateRect(hWnd, nullptr, FALSE);
                }
                return 0; // カーソル処理を行わずに終了
            }

            int clickPos = GetCursorPosFromMouse(hWnd, xPos, yPos);
            if ((GetKeyState(VK_SHIFT) & 0x8000) == 0) {
                int selStart = std::min(g_cursorPos, g_selectionStart);
                int selEnd = std::max(g_cursorPos, g_selectionStart);
                if (selStart != selEnd && clickPos >= selStart && clickPos <= selEnd) {
                    g_cursorPos = clickPos;
                    g_selectionStart = clickPos;
                }
                else if (g_inputText.empty()) {
                    g_cursorPos = 0;
                    g_selectionStart = 0;
                }
                else {
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
                }
            }
            else {
                g_cursorPos = clickPos;
            }
            SetCapture(hWnd);
            g_isDragging = true;
            g_caretVisible = true;
            InvalidateRect(hWnd, nullptr, FALSE);
        }
        break;
    case WM_MOUSEMOVE:
    {
        int yPos = GET_Y_LPARAM(lParam);
        if (g_isDraggingScrollbar) {
            RECT rc; GetClientRect(hWnd, &rc);
            float viewHeight = rc.bottom - INPUT_AREA_HEIGHT;
            float contentHeight = g_savedFormulas.size() * LIST_ITEM_HEIGHT;
            float thumbRatio = viewHeight / contentHeight;
            float thumbHeight = std::max(30.0f, viewHeight * thumbRatio);

            float deltaY = yPos - g_dragStartY;
            float scrollDelta = deltaY * (g_maxScrollY / (viewHeight - thumbHeight));
            g_scrollOffsetY = g_dragStartScrollY + scrollDelta;
            if (g_scrollOffsetY < 0) g_scrollOffsetY = 0;
            if (g_scrollOffsetY > g_maxScrollY) g_scrollOffsetY = g_maxScrollY;
            InvalidateRect(hWnd, nullptr, FALSE);
            return 0;
        }

        if (g_isDragging) {
            int xPos = GET_X_LPARAM(lParam);
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
        if (g_isDraggingScrollbar) {
            g_isDraggingScrollbar = false;
            ReleaseCapture();
            return 0;
        }
        if (g_isDragging) {
            ReleaseCapture();
            g_isDragging = false;
        }
    }
    break;
    case WM_MOUSEWHEEL:
    {
        float delta = GET_WHEEL_DELTA_WPARAM(wParam);
        g_scrollOffsetY -= (delta / 120.0f) * 40.0f;
        if (g_scrollOffsetY < 0) g_scrollOffsetY = 0;
        if (g_scrollOffsetY > g_maxScrollY) g_scrollOffsetY = g_maxScrollY;
        InvalidateRect(hWnd, nullptr, FALSE);
    }
    break;
    case WM_KEYDOWN:
    {
        bool shiftDown = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        RECT rc; GetClientRect(hWnd, &rc);
        float viewHeight = rc.bottom - INPUT_AREA_HEIGHT;

        // クリップボードからのペースト処理
        auto executePaste = [&]() {
            std::wstring pastedText = PasteFromClipboard(hWnd);
            if (!pastedText.empty()) {
                if (g_cursorPos != g_selectionStart) DeleteSelection();
                g_inputText.insert(g_cursorPos, pastedText);
                g_cursorPos += static_cast<int>(pastedText.length());
                g_selectionStart = g_cursorPos;
            }
            };

        // クリップボードにコピーする前のテキスト整形処理
        auto cleanForClipboard = [](const std::wstring& text) {
            std::wstring res;
            for (wchar_t c : text) {
                if (c == L'\x02' || c == L'\x200B') continue;
                if (c == L'\x00A0') { res += L' '; continue; }
                res += c;
            }
            return res;
            };

        switch (wParam) {
        case VK_UP:
            if (!g_savedFormulas.empty()) {
                if (g_selectedIndex > 0) {
                    g_selectedIndex--;
                }
                else if (g_selectedIndex == -1) {
                    g_selectedIndex = (int)g_savedFormulas.size() - 1;
                }
                g_inputText = g_savedFormulas[g_selectedIndex].formula;
                g_resultText = g_savedFormulas[g_selectedIndex].result;
                g_cursorPos = (int)g_inputText.length();
                g_selectionStart = g_cursorPos;

                float itemTop = g_selectedIndex * LIST_ITEM_HEIGHT;
                if (g_scrollOffsetY > itemTop) g_scrollOffsetY = itemTop;

                g_pLayoutTree.reset();
                g_pResultTree.reset();
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            return 0; // 変更: break ではなく return 0; にして再計算を回避

        case VK_DOWN:
            if (!g_savedFormulas.empty()) {
                if (g_selectedIndex < (int)g_savedFormulas.size() - 1) {
                    g_selectedIndex++;
                }
                else if (g_selectedIndex == -1) {
                    g_selectedIndex = 0;
                }
                g_inputText = g_savedFormulas[g_selectedIndex].formula;
                g_resultText = g_savedFormulas[g_selectedIndex].result;
                g_cursorPos = (int)g_inputText.length();
                g_selectionStart = g_cursorPos;

                float itemBottom = (g_selectedIndex + 1) * LIST_ITEM_HEIGHT;
                if (g_scrollOffsetY + viewHeight < itemBottom) g_scrollOffsetY = itemBottom - viewHeight;

                g_pLayoutTree.reset();
                g_pResultTree.reset();
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            return 0; // 変更: break ではなく return 0; にして再計算を回避

        case VK_ESCAPE:
            g_selectedIndex = -1;
            InvalidateRect(hWnd, nullptr, FALSE);
            return 0; // 変更

        case VK_DELETE:
            if (g_selectedIndex != -1) {
                g_savedFormulas.erase(g_savedFormulas.begin() + g_selectedIndex);
                SaveFormulasToFile();
                g_selectedIndex = -1;

                float contentHeight = g_savedFormulas.size() * LIST_ITEM_HEIGHT;
                g_maxScrollY = std::max(0.0f, contentHeight - viewHeight);
                if (g_scrollOffsetY > g_maxScrollY) g_scrollOffsetY = g_maxScrollY;

                InvalidateRect(hWnd, nullptr, FALSE);
                return 0; // 追加: リスト項目の削除時は再計算を回避
            }
            else {
                if (g_cursorPos != g_selectionStart) {
                    DeleteSelection();
                }
                else if (g_cursorPos < (int)g_inputText.length()) {
                    g_inputText.erase(g_cursorPos, 1);
                }
            }
            break; // 通常の文字削除時は break して下部の再計算を通す

        case VK_RETURN:
            g_resultText = CalculateResult(g_inputText);
            if (!g_resultText.empty()) {
                CopyToClipboard(hWnd, cleanForClipboard(g_resultText));
            }
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
            if (ctrlDown) {
                if (g_cursorPos != g_selectionStart) {
                    int startIdx = std::min(g_cursorPos, g_selectionStart);
                    int endIdx = std::max(g_cursorPos, g_selectionStart);
                    CopyToClipboard(hWnd, cleanForClipboard(g_inputText.substr(startIdx, endIdx - startIdx)));
                }
                else if (!g_inputText.empty()) {
                    std::wstring formula = cleanForClipboard(g_inputText);
                    if (!g_resultText.empty()) {
                        std::wstring ans = cleanForClipboard(g_resultText);
                        std::wstring copyText;
                        if (ans == L"is prime") {
                            copyText = formula + L" " + ans;
                        }
                        else {
                            copyText = formula + L" = " + ans;
                        }
                        CopyToClipboard(hWnd, copyText);
                    }
                    else {
                        CopyToClipboard(hWnd, formula);
                    }
                }
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
        g_selectedIndex = -1;
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
                        prevChar == L' ') {
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
        D2D1_COLOR_F bgColor = g_isDarkMode ? D2D1::ColorF(0x1E1E1E) : D2D1::ColorF(D2D1::ColorF::White);
        g_pRenderTarget->Clear(bgColor);
        Direct2DContext ctx(g_pRenderTarget.Get(), g_pDWriteFactory.Get());
        D2D1_COLOR_F textColor = g_isDarkMode ? D2D1::ColorF(0xD4D4D4) : D2D1::ColorF(D2D1::ColorF::Black);
        ctx.SetTextColor(textColor.r, textColor.g, textColor.b, textColor.a);

        float startX = 50.0f;
        float startY = INPUT_AREA_HEIGHT / 2.0f;

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
                y1 = y2;
                scale1 = scale2;
                if (std::abs(x1 - x2) > 0.01f) {
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

        RECT rc;
        GetClientRect(hWnd, &rc);

        if (!g_resultText.empty()) {
            if (g_isDarkMode) {
                ctx.SetTextColor(0.5f, 0.5f, 0.5f);
            }
            else {
                ctx.SetTextColor(0.6f, 0.6f, 0.6f);
            }
            float gap = 20.0f;
            float resultX = startX + g_pLayoutTree->width + gap;
            std::wstring cleanInput = g_inputText;
            cleanInput.erase(std::remove(cleanInput.begin(), cleanInput.end(), L'\x200B'), cleanInput.end());
            cleanInput.erase(std::remove(cleanInput.begin(), cleanInput.end(), L' '), cleanInput.end());
            if (cleanInput != L"version" && g_resultText != L"is\x00A0prime") {
                float eqW, eqH, eqD;
                ctx.MeasureGlyph(L"=", 36.0f, false, eqW, eqH, eqD);
                ctx.DrawGlyph(L"=", resultX, startY, 36.0f, false);
                resultX += eqW + gap;
            }

            float btnWidth = 50.0f;
            float maxResultWidth = rc.right - resultX - btnWidth - 40.0f; // PINボタンの幅を考慮

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

            // PINボタン描画
            float btnHeight = 26.0f;
            g_pinButtonRect = D2D1::RectF(rc.right - btnWidth - 25.0f, startY - btnHeight / 2.0f, rc.right - 25.0f, startY + btnHeight / 2.0f);
            ComPtr<ID2D1SolidColorBrush> btnBrush;
            g_pRenderTarget->CreateSolidColorBrush(g_isDarkMode ? D2D1::ColorF(0x444444) : D2D1::ColorF(0xE0E0E0), &btnBrush);
            D2D1_ROUNDED_RECT rr = D2D1::RoundedRect(g_pinButtonRect, 4.0f, 4.0f);
            g_pRenderTarget->FillRoundedRectangle(&rr, btnBrush.Get());

            ComPtr<ID2D1SolidColorBrush> btnTextBrush;
            g_pRenderTarget->CreateSolidColorBrush(g_isDarkMode ? D2D1::ColorF(0xFFFFFF) : D2D1::ColorF(0x000000), &btnTextBrush);
            g_pRenderTarget->DrawTextW(L"PIN", 3, g_pUiTextFormat.Get(), g_pinButtonRect, btnTextBrush.Get());

            ctx.SetTextColor(textColor.r, textColor.g, textColor.b, textColor.a);
        }
        else {
            g_pinButtonRect = { 0,0,0,0 };
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

        // 境界線
        ctx.SetTextColor(0.5f, 0.5f, 0.5f, 0.3f);
        ctx.DrawLine(20.0f, INPUT_AREA_HEIGHT, rc.right - 20.0f, INPUT_AREA_HEIGHT, 1.0f);

        // 保存リスト表示領域
        D2D1_RECT_F listClipRect = D2D1::RectF(0.0f, INPUT_AREA_HEIGHT, (float)rc.right, (float)rc.bottom);
        g_pRenderTarget->PushAxisAlignedClip(listClipRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        float viewHeight = rc.bottom - INPUT_AREA_HEIGHT;
        float contentHeight = g_savedFormulas.size() * LIST_ITEM_HEIGHT;
        g_maxScrollY = std::max(0.0f, contentHeight - viewHeight);
        if (g_scrollOffsetY > g_maxScrollY) g_scrollOffsetY = g_maxScrollY;
        if (g_scrollOffsetY < 0.0f) g_scrollOffsetY = 0.0f;

        float listY = INPUT_AREA_HEIGHT - g_scrollOffsetY + (LIST_ITEM_HEIGHT / 2.0f);

        ctx.SetTextColor(textColor.r, textColor.g, textColor.b, 0.8f);
        int i = 0;
        for (auto& sf : g_savedFormulas) {
            if (listY + LIST_ITEM_HEIGHT / 2.0f < INPUT_AREA_HEIGHT || listY - LIST_ITEM_HEIGHT / 2.0f > rc.bottom) {
                listY += LIST_ITEM_HEIGHT;
                continue;
            }

            if (i == g_selectedIndex) {
                ComPtr<ID2D1SolidColorBrush> highlightBrush;
                g_pRenderTarget->CreateSolidColorBrush(
                    g_isDarkMode ? D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.05f) : D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.05f),
                    &highlightBrush
                );
                D2D1_RECT_F rowRect = D2D1::RectF(0, listY - LIST_ITEM_HEIGHT / 2.0f, (float)rc.right, listY + LIST_ITEM_HEIGHT / 2.0f);
                g_pRenderTarget->FillRectangle(rowRect, highlightBrush.Get());
            }

            if (!sf.pFormulaTree) sf.pFormulaTree = ParseMathText(sf.formula, &ctx);
            if (!sf.pResultTree) sf.pResultTree = ParseMathText(sf.result, &ctx);

            float itemStartX = 40.0f;
            sf.pFormulaTree->Draw(&ctx, itemStartX, listY);

            float gap = 20.0f;
            float resX = itemStartX + sf.pFormulaTree->width + gap;
            float eqW, eqH, eqD;
            ctx.MeasureGlyph(L"=", 36.0f, false, eqW, eqH, eqD);
            ctx.DrawGlyph(L"=", resX, listY, 36.0f, false);
            resX += eqW + gap;

            sf.pResultTree->Draw(&ctx, resX, listY);

            // --- ここから追加 ---
            // 削除ボタン(×)の描画
            float delBtnWidth = 24.0f;
            float delBtnHeight = 24.0f;
            D2D1_RECT_F delBtnRect = D2D1::RectF((float)rc.right - 60.0f, listY - delBtnHeight / 2.0f, (float)rc.right - 60.0f + delBtnWidth, listY + delBtnHeight / 2.0f);

            ComPtr<ID2D1SolidColorBrush> delBrush;
            g_pRenderTarget->CreateSolidColorBrush(g_isDarkMode ? D2D1::ColorF(0x662222) : D2D1::ColorF(0xFFCCCC), &delBrush);
            D2D1_ROUNDED_RECT drr = D2D1::RoundedRect(delBtnRect, 4.0f, 4.0f);
            g_pRenderTarget->FillRoundedRectangle(&drr, delBrush.Get());

            ComPtr<ID2D1SolidColorBrush> delTextBrush;
            g_pRenderTarget->CreateSolidColorBrush(g_isDarkMode ? D2D1::ColorF(0xFF9999) : D2D1::ColorF(0xCC0000), &delTextBrush);
            g_pRenderTarget->DrawTextW(L"×", 1, g_pUiTextFormat.Get(), delBtnRect, delTextBrush.Get());
            // --- ここまで追加 ---

            ctx.SetTextColor(0.5f, 0.5f, 0.5f, 0.1f);
            ctx.DrawLine(40.0f, listY + LIST_ITEM_HEIGHT / 2.0f, rc.right - 40.0f, listY + LIST_ITEM_HEIGHT / 2.0f, 1.0f);
            ctx.SetTextColor(textColor.r, textColor.g, textColor.b, 0.8f);

            listY += LIST_ITEM_HEIGHT;
        }

        g_pRenderTarget->PopAxisAlignedClip();

        // カスタムスクロールバーの描画
        if (g_maxScrollY > 0.0f) {
            float thumbRatio = viewHeight / contentHeight;
            float thumbHeight = std::max(30.0f, viewHeight * thumbRatio);
            float thumbY = INPUT_AREA_HEIGHT + (g_scrollOffsetY / g_maxScrollY) * (viewHeight - thumbHeight);
            g_scrollbarThumbRect = D2D1::RectF((float)rc.right - 12.0f, thumbY, (float)rc.right - 4.0f, thumbY + thumbHeight);

            ComPtr<ID2D1SolidColorBrush> scrollBrush;
            g_pRenderTarget->CreateSolidColorBrush(g_isDarkMode ? D2D1::ColorF(0.4f, 0.4f, 0.4f, 0.8f) : D2D1::ColorF(0.6f, 0.6f, 0.6f, 0.8f), &scrollBrush);
            D2D1_ROUNDED_RECT srr = D2D1::RoundedRect(g_scrollbarThumbRect, 4.0f, 4.0f);
            g_pRenderTarget->FillRoundedRectangle(&srr, scrollBrush.Get());
        }
        else {
            g_scrollbarThumbRect = { 0,0,0,0 };
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
        if (g_isDemoMode) {
            KillTimer(hWnd, DEMO_TYPING_TIMER_ID);
            KillTimer(hWnd, DEMO_PAUSE_TIMER_ID);
        }
        PostQuitMessage(0);
        break;
    case WM_ERASEBKGND:
        return 1;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow) {
    std::wstring cmdLine(lpCmdLine);
    if (false) {
        g_isDemoMode = true;
        InitDemoFormulas();
    }
    D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, g_pD2DFactory.GetAddressOf());
    DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory3), reinterpret_cast<IUnknown**>(g_pDWriteFactory.GetAddressOf()));

    g_pDWriteFactory->CreateTextFormat(L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"en-us", &g_pUiTextFormat);
    g_pUiTextFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
    g_pUiTextFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

    WNDCLASSEXW wcex = { sizeof(WNDCLASSEX) };
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.lpszClassName = L"4mulaWindow";
    RegisterClassExW(&wcex);
    HWND hWnd = CreateWindowW(L"4mulaWindow", L"4mula", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, 800, 480, nullptr, nullptr, hInstance, nullptr);
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