using System;
using NAudio.CoreAudioApi;
using mute_it;

class NotificationClient: NAudio.CoreAudioApi.Interfaces.IMMNotificationClient
{
    private MuteItContext context;

    public NotificationClient(MuteItContext context)
    {
        if (Environment.OSVersion.Version.Major < 6)
        {
            throw new NotSupportedException("This feature requires Windows Vista or newer");
        }
        this.context = context;
    }

    public void OnDefaultDeviceChanged(DataFlow dataFlow, Role deviceRole, string defaultDeviceId)
    {
        context.updatePrimaryMicDevice();
    }

    public void OnDeviceAdded(string deviceId)
    {
        // Nothing to do
    }

    public void OnDeviceRemoved(string deviceId)
    {
        // Nothing to do
    }

    public void OnDeviceStateChanged(string deviceId, DeviceState newState)
    {
        // Nothing to do
    }

    public void OnPropertyValueChanged(string deviceId, PropertyKey propertyKey)
    {
        // Nothing to do
    }
}
