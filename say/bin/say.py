"""
Say for Splunk by Marinus van Aswegen (mvanaswegen AT gmail.com)

Allow Splunk to say stuff

Examples

  | say field=name 
  | say field=name preamble="The naughty user is" prelude="lock his account"
  | say field=name preamble="Interesting" mention=true
  | say field=name intro="the following people are responsible"
  | say field=error max_words=80 max_sentences=20
  
intro, intro's the search results
peramble, adds words for each field in the search result
prelude, adds words after each field in the search result  
mention, will say the name of the field
max_words, truncates a field
max_sentences, truncates a search result
  
Copyright 2011 Marinus van Aswegen. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are
permitted provided that the following conditions are met:

   1. Redistributions of source code must retain the above copyright notice, this list of
      conditions and the following disclaimer.

   2. Redistributions in binary form must reproduce the above copyright notice, this list
      of conditions and the following disclaimer in the documentation and/or other materials
      provided with the distribution.

THIS SOFTWARE IS PROVIDED BY MARINUS VAN ASWEGEN ``AS IS'' AND ANY EXPRESS OR IMPLIED
WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL MARINUS VAN ASWEGEN OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

The views and conclusions contained in the software and documentation are those of the
authors and should not be interpreted as representing official policies, either expressed
or implied, of Marinus van Aswegen.

"""

import splunk.Intersplunk
import sys
import time
import platform
import os

max_words = 40		# truncate if more than about 40 words
max_sentences = 10	# truncate number of sentences per batch
preamble = ''
prelude = ''
	
def say(words):
	if 'Windows' in platform.uname()[0]:
		from win32com.client import constants
		import win32com.client
	
		speaker = win32com.client.Dispatch("SAPI.SpVoice")
		
		# trancate words max_words * average nr of letters + space
		speaker.Speak(words[:int(max_words)*4])
		del speaker
	elif 'Darwin' in platform.uname()[0]:
		# trancate words max_words * average nr of letters + space
		cmd = 'say "%s"' % words[:int(max_words)*4]
		os.system(cmd)

try:   
	keywords,options = splunk.Intersplunk.getKeywordsAndOptions()

	
	if options.has_key('debug'):
		say(options['debug'])
		splunk.Intersplunk.generateErrorResults("saying " + options['debug'])
		exit(0)
	
	if not options.has_key('field'):
		splunk.Intersplunk.generateErrorResults("no field specified")
		exit(0)

	max_words = options.get('max_words', max_words)
	prelude = options.get('prelude', prelude)
	preamble = options.get('preamble', preamble)
	field = options.get('field', None)
	mention = options.get('mention', '')
	
	if 'true' in mention.lower():
		mention = True
	else:
		mention = False
	
	max_sentences = options.get('max_sentences', max_sentences)
	intro = options.get('intro', '')
	
	# get the previous search results
	results,unused1,unused2 = splunk.Intersplunk.getOrganizedResults()
	
	# intro the results
	if intro:
		say(intro)
		
	counter = 0
	for result in results:
		if result.has_key(field):
			if counter <= int(max_sentences): # don't over whelm us
				if mention: # add the name of the field into the sentence
					sentence = prelude + ' ' + 'field name ' + field + ' is ' + result[field] + preamble
				else:
					sentence = prelude + ' ' + result[field] + preamble
				say(sentence)
				counter += 1

	# output results
	splunk.Intersplunk.outputResults(results)


except Exception, e:
	results = splunk.Intersplunk.generateErrorResults(str(e))


