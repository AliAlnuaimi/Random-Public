using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class FPSCounter : MonoBehaviour
{
    [SerializeField] private float accum = 0.0f;
    [SerializeField] private int frames = 0;
    [SerializeField] private float timeLeft;
    [SerializeField] private float updateInterval = 0.5f;
    [SerializeField] TMPro.TextMeshProUGUI fpsCounterText;

    private void Update()
    {
        UpdateFPSCounter();
        UpdateFPSCounterText();
    }

    float fps = 0.0f;
    private void UpdateFPSCounter()
    {
        timeLeft -= Time.deltaTime;
        accum += Time.timeScale / Time.deltaTime;
        frames++;

        if (timeLeft <= 0.0f)
        {
            fps = accum / frames;
            timeLeft = updateInterval;
            accum = 0.0f;
            frames = 0;
        }
    }

    private void UpdateFPSCounterText()
    {
        string fpsCountString = string.Format("{0:F2} FPS", fps);
        fpsCounterText.text = fpsCountString;
    }
}
