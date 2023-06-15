using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;


namespace Utilities
{
    /// <summary>
    /// <para>Scroll Wheel Event Handler</para>
    /// <para>A representation of a scroll wheel input device for interacting with a web page.</para>
    /// <para>https://www.selenium.dev/documentation/webdriver/actions_api/wheel/</para>
    /// </summary>
    public static class ScrollWheelEventHandler
    {
        public enum Direction { Up, Right, Down, Left }
        // ΔX: (+) Right  (-) Left
        // ΔY: (+) Down   (-) Up

        /// <summary>
        /// <para>
        /// This actions class does not automatically scroll the target element into view, so this
        /// method will need to be used if elements are not already inside the viewport.
        /// </para>
        /// <para>
        /// Regardless of whether the element is above or below the current viewscreen, the viewport will
        /// be scrolled so the bottom of the element is at the bottom of the screen.
        /// </para>
        /// </summary>
        /// <param name="driver">Web driver</param>
        /// <param name="by">Target element</param>
        /// <param name="delay">Delay in ms</param>
        public static void Scroll_To_Element(IWebDriver driver, By by, int delay = 2000)
        {
            try
            {
                IWebElement iframe = driver.FindElement(by);

                new Actions(driver)
                    .ScrollToElement(iframe)
                    .Perform();

                Thread.Sleep(delay);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught! " + ex.Message);
            }
        }

        /// <summary>
        /// <para>
        /// Pass in an delta x and a delta y value for how much to scroll in the right and down
        /// directions. Negative values represent left and up, respectively.
        /// </para>
        /// </summary>
        /// <param name="driver">Web driver</param>
        /// <param name="direction">Enum direction of scroll</param>
        /// <param name="amount">How much to scroll</param>
        /// <param name="delay">Delay in ms</param>
        public static void Scroll_By_Given_Amount(IWebDriver driver, Direction direction = Direction.Down, int amount = 0, int delay = 2000)
        {
            try
            {
                //IWebElement iframe = driver.FindElement(by);
                (int deltaX, int deltaY) = GetDeltasFromDirection(direction, amount);

                new Actions(driver)
                    .ScrollByAmount(deltaX, deltaY)
                    .Perform();

                Thread.Sleep(delay);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught! " + ex.Message);
            }
        }

        /// <summary>
        /// <para>
        /// To execute this use the “Scroll From” method, which takes 3 arguments.The first
        /// represents the origination point, which we designate as the element, and the second two
        /// are the delta x and delta y values.
        /// </para>
        /// <para>
        /// If the element is out of the viewport, it will be scrolled to the bottom of the screen,
        /// then the page will be scrolled by the provided delta x and delta y values.
        /// </para>
        /// </summary>
        /// <param name="driver">Web driver</param>
        /// <param name="by">Target element</param>
        /// <param name="direction">Enum direction of scroll</param>
        /// <param name="amount">How much to scroll</param>
        /// <param name="delay">Delay in ms</param>
        public static void Scroll_From_An_Element_By_A_Given_Amount(IWebDriver driver, By by, Direction direction = Direction.Down, int amount = 0, int delay = 2000)
        {
            try
            {
                IWebElement iframe = driver.FindElement(by);
                WheelInputDevice.ScrollOrigin scrollOrigin = new WheelInputDevice.ScrollOrigin
                {
                    Element = iframe
                };

                (int deltaX, int deltaY) = GetDeltasFromDirection(direction, amount);

                new Actions(driver)
                    .ScrollFromOrigin(scrollOrigin, deltaX, deltaY)
                    .Perform();

                Thread.Sleep(delay);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught! " + ex.Message);
            }
        }

        /// <summary>
        /// <para>
        /// This scenario is used when you need to scroll only a portion of the screen, and it is
        /// outside the viewport. Or is inside the viewport and the portion of the screen that must
        /// be scrolled is a known offset away from a specific element.
        /// </para>
        /// <para>
        /// This uses the “Scroll From” method again, and in addition to specifying the element, an
        /// offset is specified to indicate the origin point of the scroll.The offset is calculated
        /// from the center of the provided element.
        /// </para>
        /// <para>
        /// If the element is out of the viewport, it first will be scrolled to the bottom of the
        /// screen, then the origin of the scroll will be determined by adding the offset to the
        /// coordinates of the center of the element, and finally the page will be scrolled by the
        /// provided delta x and delta y values.
        /// </para>
        /// <para>
        /// Note that if the offset from the center of the element falls outside of the viewport,
        /// it will result in an exception.
        /// </para>
        /// </summary>
        /// <param name="driver">Web driver</param>
        /// <param name="by">Target element</param>
        /// <param name="direction">Enum direction of scroll</param>
        /// <param name="amount">How much to scroll</param>
        /// <param name="XOffset">X origin point of scroll</param>
        /// <param name="YOffset">Y origin point of scroll</param>
        /// <param name="delay">Delay in ms</param>
        public static void Scroll_From_An_Element_With_An_Offset(IWebDriver driver, By by, Direction direction = Direction.Down, int amount = 0, int XOffset = 0, int YOffset = 0, int delay = 2000)
        {
            try
            {
                IWebElement iframe = driver.FindElement(by);
                var scrollOrigin = new WheelInputDevice.ScrollOrigin
                {
                    Element = iframe,
                    XOffset = XOffset,
                    YOffset = YOffset
                };

                (int deltaX, int deltaY) = GetDeltasFromDirection(direction, amount);

                new Actions(driver)
                    .ScrollFromOrigin(scrollOrigin, deltaX, deltaY)
                    .Perform();

                Thread.Sleep(delay);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught! " + ex.Message);
            }
        }

        /// <summary>
        /// <para>
        /// The final scenario is used when you need to scroll only a portion of the screen, and it
        /// is already inside the viewport.
        /// </para>
        /// <para>
        /// This uses the “Scroll From” method again, but the viewport is designated instead of an
        /// element. An offset is specified from the upper left corner of the current viewport.
        /// After the origin point is determined, the page will be scrolled by the provided delta x
        /// and delta y values.
        /// </para>
        /// <para>
        /// Note that if the offset from the upper left corner of the viewport falls outside of the
        /// screen, it will result in an exception.
        /// </para>
        /// </summary>
        /// <param name="driver">Web driver</param>
        /// <param name="direction">Enum direction of scroll</param>
        /// <param name="amount">How much to scroll</param>
        /// <param name="XOffset">X origin point of scroll</param>
        /// <param name="YOffset">Y origin point of scroll</param>
        /// <param name="delay">Delay in ms</param>
        public static void Scroll_From_A_Offset_Of_Origin_Element_By_Given_Amount(IWebDriver driver, Direction direction = Direction.Down, int amount = 0, int XOffset = 0, int YOffset = 0, int delay = 2000)
        {
            try
            {
                var scrollOrigin = new WheelInputDevice.ScrollOrigin
                {
                    Viewport = true,
                    XOffset = XOffset,
                    YOffset = YOffset
                };

                (int deltaX, int deltaY) = GetDeltasFromDirection(direction, amount);

                new Actions(driver)
                    .ScrollFromOrigin(scrollOrigin, deltaX, deltaY)
                    .Perform();

                Thread.Sleep(delay);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught! " + ex.Message);
            }
        }

        private static (int, int) GetDeltasFromDirection(Direction direction, int amount)
        {
            int deltaX = 0; // + R - L
            int deltaY = 0; // + D - U
            switch (direction)
            {
                case Direction.Up: { deltaY = -1 * amount; break; }
                case Direction.Right: { deltaX = amount; break; }
                case Direction.Down: { deltaY = amount; break; }
                case Direction.Left: { deltaX = -1 * amount; break; }
            }
            return (deltaX, deltaY);
        }
    }
}