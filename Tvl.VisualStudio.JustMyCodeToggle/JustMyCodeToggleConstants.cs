﻿// Copyright (c) Tunnel Vision Laboratories, LLC. All Rights Reserved.
// Licensed under the MIT License. See LICENSE.txt in the project root for license information.

namespace Tvl.VisualStudio.JustMyCodeToggle
{
    using Guid = System.Guid;

    internal static class JustMyCodeToggleConstants
    {
        public const string GuidJustMyCodeTogglePackageString = "73702199-D0C5-4863-8203-99D41B34DD2D";

        public const string GuidJustMyCodeToggleCommandSetString = "B1317E30-CFDA-47CA-90A0-95E894B150A0";
        public static readonly Guid GuidJustMyCodeToggleCommandSet = new Guid("{" + GuidJustMyCodeToggleCommandSetString + "}");

        public static readonly int CmdidJustMyCodeToggle = 0x0100;

        public static readonly int CmdidDiagnosticBuildLogToggle = 0x0200;
    }
}
