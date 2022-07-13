/*
 * RateStatic.cs
 *
 * authors: Tomasz Rybiarczyk (tomasz.rybiarczyk@steelseries.com)
 *
 *
 * Copyright (c) 2016 SteelSeries
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
 * CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
 * TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

namespace SteelSeries {

    namespace GameSense {

        [FullSerializer.fsObject(Converter = typeof(RateStaticConverter))]
        public class RateStatic : AbstractRate {

            public System.UInt32 frequency;
            public System.UInt32 repeatLimit;  // required but can specify 0 for no limit

            private static RateStatic _new() {
                RateStatic rs = new();
                return rs;
            }

            public static RateStatic Create( System.UInt32 frequency, System.UInt32 repeatLimit = 0 ) {
                RateStatic rs = _new();
                rs.frequency = frequency; rs.repeatLimit = repeatLimit;
                return rs;
            }

        }

        class RateStaticConverter : Converter< RateStatic > {
            protected override FullSerializer.fsResult DoSerialize( RateStatic model, System.Collections.Generic.Dictionary< string, FullSerializer.fsData > serialized ) {

                SerializeMember( serialized, null, "frequency", model.frequency );
                SerializeMember( serialized, null, "repeat_limit", model.repeatLimit );

                return FullSerializer.fsResult.Success;
            }
        }

    }

}