import React, { useState } from 'react';
import Papa from 'papaparse';

// Sample data for testing
const SAMPLE_COURSES = [
  {
    id: 1,
    course_name: "Introduction to Project Management",
    outline: "This course covers project lifecycle, planning, execution, monitoring and control.",
    objectives: "Understand project management fundamentals, create project plans, identify risks.",
    competencies: "Project planning, risk management, team leadership",
    why_attend: "Gain essential project management skills for successful projects."
  },
  {
    id: 2,
    course_name: "Advanced Project Management",
    outline: "Advanced techniques including agile methodologies and complex resource allocation.",
    objectives: "Master advanced project methodologies and sophisticated risk management.",
    competencies: "Agile project management, earned value analysis, recovery planning",
    why_attend: "Take your project management skills to the next level."
  },
  {
    id: 3,
    course_name: "Business Analysis Fundamentals",
    outline: "Introduction to business analysis including requirements elicitation and documentation.",
    objectives: "Develop requirements elicitation skills and create effective documentation.",
    competencies: "Requirements gathering, stakeholder analysis, process modeling",
    why_attend: "Learn to identify business needs and recommend effective solutions."
  }
];

const CourseComparisonTool = () => {
  const [courses, setCourses] = useState(SAMPLE_COURSES);
  const [course1, setCourse1] = useState(null);
  const [course2, setCourse2] = useState(null);
  const [results, setResults] = useState(null);
  const [fileUploaded, setFileUploaded] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  
  // Dynamic weights and thresholds
  const [weights, setWeights] = useState({
    name: 20,
    outline: 25,
    objectives: 25,
    whyAttend: 25,
    competencies: 5
  });
  
  const [thresholds, setThresholds] = useState({
    high: 75,
    medium: 50,
    low: 30
  });
  
  const [weekGaps, setWeekGaps] = useState({
    high: 4,
    medium: 3,
    low: 2
  });
  
  // Search for course
  const [search1, setSearch1] = useState('');
  const [search2, setSearch2] = useState('');
  
  // Handle file upload
  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    
    setIsLoading(true);
    setError('');
    
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.data && results.data.length > 0) {
          // Clean the data
          const cleanedData = results.data
            .filter(row => {
              // Check if row has at least a course name or title
              return Object.values(row).some(val => val && val.toString().trim() !== '');
            })
            .map((row, index) => {
              // Clean up row data and ensure it has an id
              const cleanRow = { id: index + 1 };
              
              Object.entries(row).forEach(([key, value]) => {
                // Clean keys (convert to snake_case)
                const cleanKey = key.trim().toLowerCase().replace(/\s+/g, '_');
                // Clean values (trim and remove HTML)
                const cleanValue = value && typeof value === 'string' 
                  ? value.trim().replace(/<[^>]*>/g, '') 
                  : value;
                
                cleanRow[cleanKey] = cleanValue;
              });
              
              return cleanRow;
            });
          
          setCourses(cleanedData);
          setFileUploaded(true);
          setIsLoading(false);
          setCourse1(null);
          setCourse2(null);
          setResults(null);
        } else {
          setError('The uploaded file contains no valid data');
          setIsLoading(false);
        }
      },
      error: (error) => {
        setError(`Error parsing CSV: ${error.message}`);
        setIsLoading(false);
      }
    });
  };
  
  // Get filtered courses based on search
  const getFilteredCourses = (search) => {
    if (!search.trim()) return [];
    const searchLower = search.toLowerCase();
    return courses
      .filter(course => {
        const courseName = course.course_name || course.title || course.name || '';
        return courseName.toLowerCase().includes(searchLower);
      })
      .slice(0, 10); // Limit to 10 results
  };
  
  const filteredCourses1 = getFilteredCourses(search1);
  const filteredCourses2 = getFilteredCourses(search2);
  
  // Calculate similarity (simple word overlap)
  const calculateSimilarity = (text1, text2) => {
    if (!text1 || !text2) return 0;
    
    // Tokenize and filter out short words
    const words1 = text1.toLowerCase().split(/\W+/).filter(w => w.length > 2);
    const words2 = text2.toLowerCase().split(/\W+/).filter(w => w.length > 2);
    
    // Create sets for unique words
    const set1 = new Set(words1);
    const set2 = new Set(words2);
    
    // Find intersection
    const intersection = new Set([...set1].filter(x => set2.has(x)));
    
    // Calculate Jaccard similarity
    return (intersection.size / Math.max(1, Math.min(set1.size, set2.size))) * 100;
  };
  
  // Calculate semantic similarity (context-based)
  const calculateSemanticSimilarity = (text1, text2) => {
    if (!text1 || !text2) return 0;
    
    // Count word frequencies
    const wordFreq1 = {};
    const wordFreq2 = {};
    
    text1.toLowerCase().split(/\W+/).filter(w => w.length > 2)
      .forEach(word => { wordFreq1[word] = (wordFreq1[word] || 0) + 1; });
      
    text2.toLowerCase().split(/\W+/).filter(w => w.length > 2)
      .forEach(word => { wordFreq2[word] = (wordFreq2[word] || 0) + 1; });
    
    // Compare word importance
    let commonImportance = 0;
    let totalImportance = 0;
    
    // All unique words from both texts
    const allWords = new Set([...Object.keys(wordFreq1), ...Object.keys(wordFreq2)]);
    
    allWords.forEach(word => {
      const freq1 = wordFreq1[word] || 0;
      const freq2 = wordFreq2[word] || 0;
      
      if (freq1 > 0 && freq2 > 0) {
        // Words in both texts
        commonImportance += Math.min(freq1, freq2);
      }
      
      totalImportance += Math.max(freq1, freq2);
    });
    
    // Calculate similarity
    return totalImportance > 0 ? (commonImportance / totalImportance) * 100 : 0;
  };
  
  const getCourseField = (course, fieldName) => {
    if (!course) return '';
    
    // Check for exact match
    if (course[fieldName]) return course[fieldName];
    
    // Check for field variations
    const keys = Object.keys(course);
    
    // For course name
    if (fieldName === 'course_name') {
      const nameField = keys.find(k => 
        k === 'name' || k === 'title' || k === 'course_title' || k.includes('name')
      );
      return nameField ? course[nameField] : '';
    }
    
    // For outline
    if (fieldName === 'outline') {
      const outlineField = keys.find(k => 
        k.includes('outline') || k.includes('syllabus') || k.includes('content')
      );
      return outlineField ? course[outlineField] : '';
    }
    
    // For objectives
    if (fieldName === 'objectives') {
      const objectivesField = keys.find(k => 
        k.includes('objective') || k.includes('goal') || k.includes('learn')
      );
      return objectivesField ? course[objectivesField] : '';
    }
    
    // For competencies
    if (fieldName === 'competencies') {
      const competenciesField = keys.find(k => 
        k.includes('competenc') || k.includes('skill') || k.includes('abilit')
      );
      return competenciesField ? course[competenciesField] : '';
    }
    
    // For why attend
    if (fieldName === 'why_attend') {
      const whyAttendField = keys.find(k => 
        k.includes('why') || k.includes('attend') || k.includes('benefit')
      );
      return whyAttendField ? course[whyAttendField] : '';
    }
    
    return '';
  };
  
  const compareCourses = () => {
    if (!course1 || !course2) return;
    
    // Calculate lexical similarities
    const nameSim = calculateSimilarity(
      getCourseField(course1, 'course_name'), 
      getCourseField(course2, 'course_name')
    );
    
    const outlineSim = calculateSimilarity(
      getCourseField(course1, 'outline'), 
      getCourseField(course2, 'outline')
    );
    
    const objectivesSim = calculateSimilarity(
      getCourseField(course1, 'objectives'), 
      getCourseField(course2, 'objectives')
    );
    
    const whyAttendSim = calculateSimilarity(
      getCourseField(course1, 'why_attend'), 
      getCourseField(course2, 'why_attend')
    );
    
    const competenciesSim = calculateSimilarity(
      getCourseField(course1, 'competencies'), 
      getCourseField(course2, 'competencies')
    );
    
    // Calculate semantic similarities
    const nameSimSemantic = calculateSemanticSimilarity(
      getCourseField(course1, 'course_name'), 
      getCourseField(course2, 'course_name')
    );
    
    const outlineSimSemantic = calculateSemanticSimilarity(
      getCourseField(course1, 'outline'), 
      getCourseField(course2, 'outline')
    );
    
    const objectivesSimSemantic = calculateSemanticSimilarity(
      getCourseField(course1, 'objectives'), 
      getCourseField(course2, 'objectives')
    );
    
    const whyAttendSimSemantic = calculateSemanticSimilarity(
      getCourseField(course1, 'why_attend'), 
      getCourseField(course2, 'why_attend')
    );
    
    const competenciesSimSemantic = calculateSemanticSimilarity(
      getCourseField(course1, 'competencies'), 
      getCourseField(course2, 'competencies')
    );
    
    // Calculate weighted averages
    const totalWeight = weights.name + weights.outline + weights.objectives + 
                         weights.whyAttend + weights.competencies;
    
    const normalizedWeights = {
      name: weights.name / totalWeight,
      outline: weights.outline / totalWeight,
      objectives: weights.objectives / totalWeight,
      whyAttend: weights.whyAttend / totalWeight,
      competencies: weights.competencies / totalWeight
    };
    
    const overallLexical = 
      nameSim * normalizedWeights.name +
      outlineSim * normalizedWeights.outline + 
      objectivesSim * normalizedWeights.objectives +
      whyAttendSim * normalizedWeights.whyAttend +
      competenciesSim * normalizedWeights.competencies;
    
    const overallSemantic = 
      nameSimSemantic * normalizedWeights.name +
      outlineSimSemantic * normalizedWeights.outline + 
      objectivesSimSemantic * normalizedWeights.objectives +
      whyAttendSimSemantic * normalizedWeights.whyAttend +
      competenciesSimSemantic * normalizedWeights.competencies;
    
    // Determine scheduling recommendation based on highest similarity
    const maxSimilarity = Math.max(overallLexical, overallSemantic);
    let recommendedGap = 0;
    
    if (maxSimilarity > thresholds.high) {
      recommendedGap = weekGaps.high;
    } else if (maxSimilarity > thresholds.medium) {
      recommendedGap = weekGaps.medium;
    } else if (maxSimilarity > thresholds.low) {
      recommendedGap = weekGaps.low;
    }
    
    setResults({
      lexical: {
        name: nameSim,
        outline: outlineSim,
        objectives: objectivesSim,
        whyAttend: whyAttendSim,
        competencies: competenciesSim,
        overall: overallLexical
      },
      semantic: {
        name: nameSimSemantic,
        outline: outlineSimSemantic,
        objectives: objectivesSimSemantic,
        whyAttend: whyAttendSimSemantic,
        competencies: competenciesSimSemantic,
        overall: overallSemantic
      },
      maxSimilarity,
      recommendedGap
    });
  };

  return (
    <div className="p-4 bg-gray-50 max-w-4xl mx-auto">
      <h1 className="text-xl font-bold text-center mb-4">Course Comparison Tool</h1>
      
      {/* File Upload Section */}
      <div className="bg-white p-4 rounded shadow mb-4">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-lg font-medium">Upload Course Data</h2>
          <label className="bg-blue-600 hover:bg-blue-700 text-white py-2 px-4 rounded cursor-pointer">
            Upload CSV File
            <input
              type="file"
              accept=".csv"
              className="hidden"
              onChange={handleFileUpload}
            />
          </label>
        </div>
        
        {isLoading ? (
          <div className="text-center py-2 text-blue-600">Loading data...</div>
        ) : error ? (
          <div className="text-red-600 p-2 bg-red-50 rounded border border-red-200">{error}</div>
        ) : (
          <div className="text-gray-600 text-sm">
            {fileUploaded 
              ? `Successfully loaded ${courses.length} courses.` 
              : "Using sample data. Upload your CSV to use your own course data."}
          </div>
        )}
      </div>
      
      {/* Course Selection and Settings */}
      <div className="bg-white p-4 rounded shadow mb-4">
        <h2 className="text-lg font-medium mb-3">Select Courses to Compare</h2>
        
        {/* Course Search and Selection */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
          {/* Course 1 Selection */}
          <div>
            <label className="block text-sm font-medium mb-1">Course 1</label>
            <div className="relative">
              <input
                type="text"
                className="w-full p-2 border rounded"
                placeholder="Search for course..."
                value={search1}
                onChange={(e) => setSearch1(e.target.value)}
              />
              
              {filteredCourses1.length > 0 && (
                <div className="absolute z-10 w-full mt-1 bg-white border rounded shadow-lg max-h-48 overflow-auto">
                  {filteredCourses1.map((course) => (
                    <div 
                      key={course.id}
                      className="p-2 hover:bg-blue-50 cursor-pointer"
                      onClick={() => {
                        setCourse1(course);
                        setSearch1(getCourseField(course, 'course_name'));
                      }}
                    >
                      {getCourseField(course, 'course_name')}
                    </div>
                  ))}
                </div>
              )}
              
              {course1 && (
                <div className="mt-1 text-sm text-blue-600">
                  Selected: {getCourseField(course1, 'course_name')}
                </div>
              )}
            </div>
          </div>
          
          {/* Course 2 Selection */}
          <div>
            <label className="block text-sm font-medium mb-1">Course 2</label>
            <div className="relative">
              <input
                type="text"
                className="w-full p-2 border rounded"
                placeholder="Search for course..."
                value={search2}
                onChange={(e) => setSearch2(e.target.value)}
              />
              
              {filteredCourses2.length > 0 && (
                <div className="absolute z-10 w-full mt-1 bg-white border rounded shadow-lg max-h-48 overflow-auto">
                  {filteredCourses2.map((course) => (
                    <div 
                      key={course.id}
                      className="p-2 hover:bg-blue-50 cursor-pointer"
                      onClick={() => {
                        setCourse2(course);
                        setSearch2(getCourseField(course, 'course_name'));
                      }}
                    >
                      {getCourseField(course, 'course_name')}
                    </div>
                  ))}
                </div>
              )}
              
              {course2 && (
                <div className="mt-1 text-sm text-blue-600">
                  Selected: {getCourseField(course2, 'course_name')}
                </div>
              )}
            </div>
          </div>
        </div>
        
        {/* Comparison Settings */}
        <div className="mb-4 p-3 bg-gray-50 rounded border">
          <h3 className="text-sm font-medium mb-3">Comparison Settings</h3>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <h4 className="text-xs font-medium mb-2">Field Weights (%)</h4>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-xs mb-1">Course Name</label>
                  <input 
                    type="number" 
                    min="0" 
                    max="100"
                    className="w-full p-1 border rounded text-sm"
                    value={weights.name}
                    onChange={(e) => setWeights({...weights, name: parseInt(e.target.value) || 0})}
                  />
                </div>
                <div>
                  <label className="block text-xs mb-1">Outline</label>
                  <input 
                    type="number" 
                    min="0" 
                    max="100"
                    className="w-full p-1 border rounded text-sm"
                    value={weights.outline}
                    onChange={(e) => setWeights({...weights, outline: parseInt(e.target.value) || 0})}
                  />
                </div>
                <div>
                  <label className="block text-xs mb-1">Objectives</label>
                  <input 
                    type="number" 
                    min="0" 
                    max="100"
                    className="w-full p-1 border rounded text-sm"
                    value={weights.objectives}
                    onChange={(e) => setWeights({...weights, objectives: parseInt(e.target.value) || 0})}
                  />
                </div>
                <div>
                  <label className="block text-xs mb-1">Why Attend</label>
                  <input 
                    type="number" 
                    min="0" 
                    max="100"
                    className="w-full p-1 border rounded text-sm"
                    value={weights.whyAttend}
                    onChange={(e) => setWeights({...weights, whyAttend: parseInt(e.target.value) || 0})}
                  />
                </div>
                <div>
                  <label className="block text-xs mb-1">Competencies</label>
                  <input 
                    type="number" 
                    min="0" 
                    max="100"
                    className="w-full p-1 border rounded text-sm"
                    value={weights.competencies}
                    onChange={(e) => setWeights({...weights, competencies: parseInt(e.target.value) || 0})}
                  />
                </div>
              </div>
            </div>
            
            <div>
              <div className="mb-3">
                <h4 className="text-xs font-medium mb-2">Similarity Thresholds (%)</h4>
                <div className="grid grid-cols-3 gap-2">
                  <div>
                    <label className="block text-xs mb-1">High</label>
                    <input 
                      type="number" 
                      min="0" 
                      max="100"
                      className="w-full p-1 border rounded text-sm"
                      value={thresholds.high}
                      onChange={(e) => setThresholds({...thresholds, high: parseInt(e.target.value) || 0})}
                    />
                  </div>
                  <div>
                    <label className="block text-xs mb-1">Medium</label>
                    <input 
                      type="number" 
                      min="0" 
                      max="100"
                      className="w-full p-1 border rounded text-sm"
                      value={thresholds.medium}
                      onChange={(e) => setThresholds({...thresholds, medium: parseInt(e.target.value) || 0})}
                    />
                  </div>
                  <div>
                    <label className="block text-xs mb-1">Low</label>
                    <input 
                      type="number" 
                      min="0" 
                      max="100"
                      className="w-full p-1 border rounded text-sm"
                      value={thresholds.low}
                      onChange={(e) => setThresholds({...thresholds, low: parseInt(e.target.value) || 0})}
                    />
                  </div>
                </div>
              </div>
              
              <div>
                <h4 className="text-xs font-medium mb-2">Scheduling Gaps (weeks)</h4>
                <div className="grid grid-cols-3 gap-2">
                  <div>
                    <label className="block text-xs mb-1">High</label>
                    <input 
                      type="number" 
                      min="0"
                      className="w-full p-1 border rounded text-sm"
                      value={weekGaps.high}
                      onChange={(e) => setWeekGaps({...weekGaps, high: parseInt(e.target.value) || 0})}
                    />
                  </div>
                  <div>
                    <label className="block text-xs mb-1">Medium</label>
                    <input 
                      type="number" 
                      min="0"
                      className="w-full p-1 border rounded text-sm"
                      value={weekGaps.medium}
                      onChange={(e) => setWeekGaps({...weekGaps, medium: parseInt(e.target.value) || 0})}
                    />
                  </div>
                  <div>
                    <label className="block text-xs mb-1">Low</label>
                    <input 
                      type="number" 
                      min="0"
                      className="w-full p-1 border rounded text-sm"
                      value={weekGaps.low}
                      onChange={(e) => setWeekGaps({...weekGaps, low: parseInt(e.target.value) || 0})}
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        
        <button
          className="w-full p-2 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded disabled:bg-blue-300"
          onClick={compareCourses}
          disabled={!course1 || !course2}
        >
          Compare Courses
        </button>
      </div>
      
      {/* Comparison Results */}
      {results && (
        <div className="bg-white p-4 rounded shadow mb-4">
          <h2 className="text-lg font-medium mb-4">Comparison Results</h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
            {/* Lexical Comparison */}
            <div>
              <h3 className="text-md font-medium text-blue-800 mb-2">Lexical Comparison</h3>
              <p className="text-xs text-gray-600 mb-3">Word overlap analysis</p>
              
              <div className="space-y-3">
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Course Name</span>
                    <span>{results.lexical.name.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-blue-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.lexical.name))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Outline</span>
                    <span>{results.lexical.outline.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-blue-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.lexical.outline))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Objectives</span>
                    <span>{results.lexical.objectives.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-blue-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.lexical.objectives))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Why Attend</span>
                    <span>{results.lexical.whyAttend.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-blue-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.lexical.whyAttend))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Competencies</span>
                    <span>{results.lexical.competencies.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-blue-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.lexical.competencies))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div className="pt-2 border-t border-gray-200 mt-2">
                  <div className="flex justify-between text-sm font-medium mb-1">
                    <span>Overall</span>
                    <span>{results.lexical.overall.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-3">
                    <div 
                      className={`h-3 rounded-full ${
                        results.lexical.overall > thresholds.high ? 'bg-red-500' : 
                        results.lexical.overall > thresholds.medium ? 'bg-orange-500' : 
                        results.lexical.overall > thresholds.low ? 'bg-yellow-500' : 
                        'bg-green-500'
                      }`}
                      style={{ width: `${Math.min(100, Math.max(0, results.lexical.overall))}%` }}
                    ></div>
                  </div>
                </div>
              </div>
            </div>
            
            {/* Semantic Comparison */}
            <div>
              <h3 className="text-md font-medium text-purple-800 mb-2">Semantic Comparison</h3>
              <p className="text-xs text-gray-600 mb-3">Context and meaning analysis</p>
              
              <div className="space-y-3">
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Course Name</span>
                    <span>{results.semantic.name.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-purple-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.semantic.name))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Outline</span>
                    <span>{results.semantic.outline.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-purple-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.semantic.outline))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Objectives</span>
                    <span>{results.semantic.objectives.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-purple-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.semantic.objectives))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Why Attend</span>
                    <span>{results.semantic.whyAttend.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-purple-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.semantic.whyAttend))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div>
                  <div className="flex justify-between text-sm mb-1">
                    <span>Competencies</span>
                    <span>{results.semantic.competencies.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div 
                      className="bg-purple-600 h-2 rounded-full" 
                      style={{ width: `${Math.min(100, Math.max(0, results.semantic.competencies))}%` }}
                    ></div>
                  </div>
                </div>
                
                <div className="pt-2 border-t border-gray-200 mt-2">
                  <div className="flex justify-between text-sm font-medium mb-1">
                    <span>Overall</span>
                    <span>{results.semantic.overall.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-3">
                    <div 
                      className={`h-3 rounded-full ${
                        results.semantic.overall > thresholds.high ? 'bg-red-500' : 
                        results.semantic.overall > thresholds.medium ? 'bg-orange-500' : 
                        results.semantic.overall > thresholds.low ? 'bg-yellow-500' : 
                        'bg-green-500'
                      }`}
                      style={{ width: `${Math.min(100, Math.max(0, results.semantic.overall))}%` }}
                    ></div>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
          <div className="bg-green-50 p-4 rounded border border-green-200">
            <h3 className="text-md font-medium text-green-800 mb-2">Scheduling Recommendation</h3>
            <p className="text-md font-medium">
              {results.recommendedGap > 0 
                ? `⚠️ Schedule these courses at least ${results.recommendedGap} week${results.recommendedGap > 1 ? 's' : ''} apart` 
                : '✓ These courses can be scheduled in the same week'
              }
            </p>
            <p className="text-xs text-gray-600 mt-2">
              Based on {results.maxSimilarity.toFixed(1)}% maximum similarity with your custom weighting and thresholds.
            </p>
          </div>
        </div>
      )}
      
      {/* Course Details */}
      {course1 && course2 && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="bg-white p-4 rounded shadow">
            <h3 className="font-medium mb-2">{getCourseField(course1, 'course_name')}</h3>
            
            <div className="space-y-3">
              {getCourseField(course1, 'outline') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Outline</div>
                  <p className="text-xs text-gray-600">{getCourseField(course1, 'outline')}</p>
                </div>
              )}
              
              {getCourseField(course1, 'objectives') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Objectives</div>
                  <p className="text-xs text-gray-600">{getCourseField(course1, 'objectives')}</p>
                </div>
              )}
              
              {getCourseField(course1, 'why_attend') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Why Attend</div>
                  <p className="text-xs text-gray-600">{getCourseField(course1, 'why_attend')}</p>
                </div>
              )}
              
              {getCourseField(course1, 'competencies') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Competencies</div>
                  <p className="text-xs text-gray-600">{getCourseField(course1, 'competencies')}</p>
                </div>
              )}
            </div>
          </div>
          
          <div className="bg-white p-4 rounded shadow">
            <h3 className="font-medium mb-2">{getCourseField(course2, 'course_name')}</h3>
            
            <div className="space-y-3">
              {getCourseField(course2, 'outline') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Outline</div>
                  <p className="text-xs text-gray-600">{getCourseField(course2, 'outline')}</p>
                </div>
              )}
              
              {getCourseField(course2, 'objectives') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Objectives</div>
                  <p className="text-xs text-gray-600">{getCourseField(course2, 'objectives')}</p>
                </div>
              )}
              
              {getCourseField(course2, 'why_attend') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Why Attend</div>
                  <p className="text-xs text-gray-600">{getCourseField(course2, 'why_attend')}</p>
                </div>
              )}
              
              {getCourseField(course2, 'competencies') && (
                <div>
                  <div className="text-sm font-medium text-gray-700">Competencies</div>
                  <p className="text-xs text-gray-600">{getCourseField(course2, 'competencies')}</p>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default CourseComparisonTool; 