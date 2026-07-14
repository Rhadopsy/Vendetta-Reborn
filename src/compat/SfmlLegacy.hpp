#ifndef VENDETTA_COMPAT_SFML_LEGACY_HPP
#define VENDETTA_COMPAT_SFML_LEGACY_HPP

#include <SFML/Graphics.hpp>

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <memory>
#include <random>
#include <stdexcept>
#include <string>
#include <type_traits>

namespace legacy
{
class Image
{
public:
    Image() : mState(std::make_shared<State>()) {}

    Image(unsigned int width, unsigned int height, const sf::Color& color = sf::Color::Black) : Image()
    {
        create(width, height, color);
    }

    Image(const Image&) = default;
    Image& operator=(const Image&) = default;

    bool loadFromFile(const std::string& filename)
    {
        mState->dirty = true;
        return mState->pixels.loadFromFile(filename);
    }

    void create(unsigned int width, unsigned int height, const sf::Color& color = sf::Color::Black)
    {
        mState->pixels.create(width, height, color);
        mState->dirty = true;
    }

    unsigned int getWidth() const { return mState->pixels.getSize().x; }
    unsigned int getHeight() const { return mState->pixels.getSize().y; }
    sf::Vector2u getSize() const { return mState->pixels.getSize(); }

    sf::Color getPixel(unsigned int x, unsigned int y) const
    {
        return mState->pixels.getPixel(x, y);
    }

    void setPixel(unsigned int x, unsigned int y, const sf::Color& color)
    {
        mState->pixels.setPixel(x, y, color);
        mState->dirty = true;
    }

    const sf::Uint8* getPixelsPtr() const { return mState->pixels.getPixelsPtr(); }

    void setSmooth(bool smooth)
    {
        mState->smooth = smooth;
        if (mState->texture.getSize().x != 0)
            mState->texture.setSmooth(smooth);
    }

    const sf::Texture& texture() const
    {
        if (mState->dirty)
        {
            if (mState->pixels.getSize().x == 0 || mState->pixels.getSize().y == 0)
            {
                sf::Image fallback;
                fallback.create(1, 1, sf::Color::Transparent);
                if (!mState->texture.loadFromImage(fallback))
                    throw std::runtime_error("Unable to create fallback SFML texture");
            }
            else if (!mState->texture.loadFromImage(mState->pixels))
            {
                throw std::runtime_error("Unable to upload SFML texture");
            }
            mState->texture.setSmooth(mState->smooth);
            mState->dirty = false;
        }
        return mState->texture;
    }

private:
    struct State
    {
        sf::Image pixels;
        sf::Texture texture;
        bool smooth{false};
        bool dirty{true};
    };

    std::shared_ptr<State> mState;
};

class Sprite : public sf::Drawable, public sf::Transformable
{
public:
    Sprite() = default;
    explicit Sprite(const Image& image) { setImage(image); }

    void setImage(const Image& image)
    {
        mImage = std::make_shared<Image>(image);
        mHasRect = false;
    }

    const Image* getImage() const { return mImage.get(); }

    void setSubRect(const sf::IntRect& legacyRect)
    {
        mTextureRect = sf::IntRect(
            legacyRect.left,
            legacyRect.top,
            legacyRect.width - legacyRect.left,
            legacyRect.height - legacyRect.top);
        mHasRect = true;
    }

    void resize(double width, double height)
    {
        const auto base = baseSize();
        if (base.x > 0.f && base.y > 0.f)
            setScale(static_cast<float>(width) / base.x, static_cast<float>(height) / base.y);
    }

    void setScaleX(double scaleX)
    {
        setScale(static_cast<float>(scaleX), getScale().y);
    }

    sf::Vector2f getSize() const
    {
        const auto base = baseSize();
        const auto scale = getScale();
        return {std::abs(base.x * scale.x), std::abs(base.y * scale.y)};
    }

    sf::Color getPixel(unsigned int x, unsigned int y) const
    {
        if (!mImage)
            return sf::Color::Transparent;
        const unsigned int offsetX = mHasRect ? static_cast<unsigned int>(std::max(0, mTextureRect.left)) : 0;
        const unsigned int offsetY = mHasRect ? static_cast<unsigned int>(std::max(0, mTextureRect.top)) : 0;
        return mImage->getPixel(offsetX + x, offsetY + y);
    }

    void setColor(const sf::Color& color) { mColor = color; }
    const sf::Color& getColor() const { return mColor; }

private:
    sf::Vector2f baseSize() const
    {
        if (mHasRect)
            return {static_cast<float>(std::abs(mTextureRect.width)), static_cast<float>(std::abs(mTextureRect.height))};
        if (mImage)
            return {static_cast<float>(mImage->getWidth()), static_cast<float>(mImage->getHeight())};
        return {0.f, 0.f};
    }

    void draw(sf::RenderTarget& target, sf::RenderStates states) const override
    {
        if (!mImage)
            return;

        mDrawable.setTexture(mImage->texture(), true);
        if (mHasRect)
            mDrawable.setTextureRect(mTextureRect);
        mDrawable.setColor(mColor);
        states.transform *= getTransform();
        target.draw(mDrawable, states);
    }

    std::shared_ptr<Image> mImage;
    sf::IntRect mTextureRect;
    bool mHasRect{false};
    sf::Color mColor{sf::Color::White};
    mutable sf::Sprite mDrawable;
};

class Text : public sf::Text
{
public:
    Text() { initialize(); }
    explicit Text(const char* text) { initialize(); setString(text); }
    explicit Text(const std::string& text) { initialize(); setString(text); }
    explicit Text(const sf::String& text) { initialize(); setString(text); }

    void setText(const std::string& text) { setString(sf::String::fromUtf8(text.begin(), text.end())); }
    void setText(const char* text) { setString(text); }
    void setText(const sf::String& text) { setString(text); }
    const sf::String& getText() const { return getString(); }
    void setSize(unsigned int size) { setCharacterSize(size); }
    unsigned int getSize() const { return getCharacterSize(); }
    void setColor(const sf::Color& color) { setFillColor(color); }
    sf::FloatRect getRect() const { return getLocalBounds(); }
    void setScaleX(double scaleX) { setScale(static_cast<float>(scaleX), getScale().y); }

private:
    static const sf::Font& defaultFont()
    {
        static const sf::Font font = [] {
            sf::Font loaded;
            if (!loaded.loadFromFile("Data/Fonts/tuffy.ttf"))
                throw std::runtime_error("Unable to load Data/Fonts/tuffy.ttf");
            return loaded;
        }();
        return font;
    }

    void initialize()
    {
        setFont(defaultFont());
    }
};

template <typename T, typename U>
std::common_type_t<T, U> random(T min, U max)
{
    using Result = std::common_type_t<T, U>;
    static thread_local std::mt19937 generator(std::random_device{}());
    if constexpr (std::is_integral_v<Result>)
        return std::uniform_int_distribution<Result>(static_cast<Result>(min), static_cast<Result>(max))(generator);
    else
        return std::uniform_real_distribution<Result>(static_cast<Result>(min), static_cast<Result>(max))(generator);
}
}

#endif
