using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using static AI_Video_Assembler.MainWindow;

namespace AI_Video_Assembler
{
    public class RenderJobSettings
    {
        // ... (ตัวแปรเดิม TrimStart, Speed, etc.)

        // [เพิ่มตัวนี้] เก็บชื่อ Encoder ที่จะใช้ (เช่น h264_nvenc, h264_amf, libx264)
        public string EncoderName { get; set; } = "libx264";
        public double TrimStart { get; set; }
        public double TrimEnd { get; set; }
        public double Speed { get; set; }
        public bool MuteAudio { get; set; }
        public double OutputWidth { get; set; }
        public double OutputHeight { get; set; }
        public List<TextLayerInfo> TextLayers { get; set; } = new List<TextLayerInfo>();

    }
}
