namespace RPAChallange
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ChromeBrowser Browser = new ChromeBrowser();
            await Browser.Run();
        }
    }
}

