namespace JsonConverter
{
    public delegate void EventHandler(object sender, params object[] args);

    class EventDispatcher
    {
        private static EventDispatcher mInstance = null;
        public static EventDispatcher instance
        {
            get
            {
                if (null == mInstance)
                    mInstance = new EventDispatcher();

                return mInstance;
            }
        }

        public event EventHandler UpdateInformation;
        public event EventHandler UpdateInformationWithFilePath;
        public event EventHandler ShowMessageBox;
        public event EventHandler CommandLineExecuteStart;
        public event EventHandler CommandLineExecuteEnd;

        public void OnUpdateInformation(string information)
            => UpdateInformation(this, information);

        public void OnUpdateInformationWithFilePath(string filePath, string infomation)
            => UpdateInformationWithFilePath(this, filePath, infomation);

        public void OnShowMessageBox(string message)
            => ShowMessageBox(this, message);

        public void OnCommandLineExecuteStart()
            => CommandLineExecuteStart(this);

        public void OnCommandLineExecuteEnd()
            => CommandLineExecuteEnd(this);
    }
}
