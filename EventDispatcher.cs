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

        public void OnUpdateInformation(string information)
            => UpdateInformation(this, information);
    }
}
