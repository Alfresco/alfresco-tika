/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.apache.tika.parser.microsoft;

import static junit.framework.Assert.assertTrue;

import java.io.InputStream;

import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.tika.TikaTest;
import org.apache.tika.io.CloseShieldInputStream;
import org.apache.tika.io.TikaInputStream;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.microsoft.ooxml.OOXMLParserTest;
import org.junit.Test;

public class OfficeParserTest extends TikaTest {

	@Test
	public void parseOfficeWord() throws Exception {
		Metadata metadata = new Metadata();
		Parser parser = new OfficeParser();

		String xml = getXML(getTestDocument("test.doc"), parser, metadata).xml;

		assertTrue(xml.contains("test"));
	}

	@Test
	public void testPoiBug61295() throws Exception {

		boolean passed = false;
		DirectoryNode root;
		InputStream stream = getTestDocument("61295.tmp");
		
		try {
			TikaInputStream tstream = TikaInputStream.cast(stream);
			
			if (tstream == null) {
				root = new NPOIFSFileSystem(new CloseShieldInputStream(stream)).getRoot();
			} else {
				final Object container = tstream.getOpenContainer();
				if (container instanceof NPOIFSFileSystem) {
					root = ((NPOIFSFileSystem) container).getRoot();
				} else if (container instanceof DirectoryNode) {
					root = (DirectoryNode) container;
				} else if (tstream.hasFile()) {
					root = new NPOIFSFileSystem(tstream.getFileChannel()).getRoot();
				} else {
					root = new NPOIFSFileSystem(new CloseShieldInputStream(tstream)).getRoot();
				}
			}

			Metadata metadata = new Metadata();
			new SummaryExtractor(metadata).parseSummaries(root);
			
			passed = true;
		} catch (ArrayIndexOutOfBoundsException e) {
			passed = true;
		} catch (Exception exception) {
			passed = false;
		} finally {
			if (stream != null) {
				try {
					stream.close();
				} catch (Exception e2) {
					stream = null;
				}
			}
		}

		assertTrue("An exception that is not ArrayIndexOutOfBoundsException has been thrown", passed);
	}

	private InputStream getTestDocument(String name) {
		return TikaInputStream.get(OOXMLParserTest.class.getResourceAsStream("/test-documents/" + name));
	}
}
