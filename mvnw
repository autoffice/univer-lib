#!/bin/sh
# ----------------------------------------------------------------------------
# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# Apache Maven Wrapper startup batch script, version 3.2.0
# ----------------------------------------------------------------------------

# If JAVA_HOME is set, use it; otherwise use java from PATH
if [ -n "$JAVA_HOME" ] ; then
  if [ -x "$JAVA_HOME/jre/sh/java" ] ; then
    JAVACMD="$JAVA_HOME/jre/sh/java"
  else
    JAVACMD="$JAVA_HOME/bin/java"
  fi
  if [ ! -x "$JAVACMD" ] ; then
    die "ERROR: JAVA_HOME is set to an invalid directory: $JAVA_HOME"
  fi
else
  JAVACMD="java"
  which java > /dev/null 2>&1 || die "ERROR: JAVA_HOME is not set and no 'java' command could be found in your PATH."
fi

# Determine the script directory
APP_HOME=$(cd "$(dirname "$0")" && pwd)

# Determine the wrapper jar location
WRAPPER_JAR="$APP_HOME/.mvn/wrapper/maven-wrapper.jar"
WRAPPER_PROPERTIES="$APP_HOME/.mvn/wrapper/maven-wrapper.properties"

# Read the distributionUrl from the properties file
DISTRIBUTION_URL=$(grep "^distributionUrl" "$WRAPPER_PROPERTIES" | cut -d'=' -f2-)

# Determine the Maven home directory
MAVEN_USER_HOME="${MAVEN_USER_HOME:-$HOME/.m2}"
WRAPPER_DIR="$MAVEN_USER_HOME/wrapper"

# Extract the distribution name from the URL
DIST_NAME=$(basename "$DISTRIBUTION_URL" .zip)
MAVEN_HOME="$WRAPPER_DIR/dists/$DIST_NAME"

# Extract version from distribution URL for fallback lookup
MAVEN_VERSION=$(echo "$DIST_NAME" | sed 's/apache-maven-//;s/-bin//')

# Try to find existing Maven installation before downloading
MAVEN_EXEC=""
if [ -d "$MAVEN_HOME" ]; then
  MAVEN_EXEC="$MAVEN_HOME/bin/mvn"
elif [ -d "$WRAPPER_DIR/dists/apache-maven-$MAVEN_VERSION" ]; then
  MAVEN_EXEC="$WRAPPER_DIR/dists/apache-maven-$MAVEN_VERSION/bin/mvn"
fi

# Download and install Maven if not already present
if [ -z "$MAVEN_EXEC" ] || [ ! -x "$MAVEN_EXEC" ]; then
  mkdir -p "$WRAPPER_DIR/dists"
  DIST_ZIP="$WRAPPER_DIR/dists/$DIST_NAME.zip"
  echo "Downloading $DISTRIBUTION_URL"
  if command -v curl > /dev/null 2>&1; then
    curl -fsSL "$DISTRIBUTION_URL" -o "$DIST_ZIP"
  elif command -v wget > /dev/null 2>&1; then
    wget -q "$DISTRIBUTION_URL" -O "$DIST_ZIP"
  else
    echo "ERROR: Neither curl nor wget is available." >&2
    exit 1
  fi
  unzip -q "$DIST_ZIP" -d "$WRAPPER_DIR/dists"
  rm -f "$DIST_ZIP"
  # Find the actual Maven executable after download
  MAVEN_EXEC=$(find "$WRAPPER_DIR/dists" -name "mvn" -path "*/bin/mvn" 2>/dev/null | grep "$DIST_NAME" | head -1)
  if [ -z "$MAVEN_EXEC" ]; then
    MAVEN_EXEC=$(find "$WRAPPER_DIR/dists" -name "mvn" -path "*/bin/mvn" 2>/dev/null | head -1)
  fi
fi

if [ -z "$MAVEN_EXEC" ]; then
  echo "ERROR: Could not find Maven executable after download." >&2
  exit 1
fi

exec "$MAVEN_EXEC" "$@"
