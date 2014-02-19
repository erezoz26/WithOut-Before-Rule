package mavenTest;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestRule;

public class myTest extends MyTestCases{
		
	@Rule
	public TestRule withoutBeforeRule = new WithoutBeforeRule();
	
	@Test
	@WithoutBefore
	public void testOne() throws Exception {
		System.out.println("i'm in testOne");
		
	}
	
	@Test
	public void testTwo() throws Exception {
		System.out.println("i'm in testTwo");
	}

}
